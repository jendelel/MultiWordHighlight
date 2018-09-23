using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;
using System.Windows.Media;

namespace MultiWordHighlight
{
    #region Format definitions

    [Export(typeof(EditorFormatDefinition))]
    [Name("MarkerFormatDefinition/HighlightWordFormatDefinition1")]
    [UserVisible(true)]
    internal class HighlightWordFormatDefinition : MarkerFormatDefinition
    {
        public HighlightWordFormatDefinition()
        {
            this.BackgroundColor = Colors.Yellow;
            this.ForegroundColor = Colors.DarkRed;
            this.DisplayName = "Highlight Word 1";
            this.ZOrder = 5;
        }
    }

    [Export(typeof(EditorFormatDefinition))]
    [Name("MarkerFormatDefinition/HighlightWordFormatDefinition2")]
    [UserVisible(true)]
    internal class HighlightWordFormatDefinition2 : MarkerFormatDefinition
    {
        public HighlightWordFormatDefinition2()
        {
            this.BackgroundColor = Colors.GreenYellow;
            this.ForegroundColor = Colors.DarkRed;
            this.DisplayName = "Highlight Word 2";
            this.ZOrder = 5;
        }
    }

    [Export(typeof(EditorFormatDefinition))]
    [Name("MarkerFormatDefinition/HighlightWordFormatDefinition3")]
    [UserVisible(true)]
    internal class HighlightWordFormatDefinition3 : MarkerFormatDefinition
    {
        public HighlightWordFormatDefinition3()
        {
            this.BackgroundColor = Colors.Gold;
            this.ForegroundColor = Colors.DarkRed;
            this.DisplayName = "Highlight Word 3";
            this.ZOrder = 5;
        }
    }

    [Export(typeof(EditorFormatDefinition))]
    [Name("MarkerFormatDefinition/HighlightWordFormatDefinition4")]
    [UserVisible(true)]
    internal class HighlightWordFormatDefinition4 : MarkerFormatDefinition
    {
        public HighlightWordFormatDefinition4()
        {
            this.BackgroundColor = Colors.Lime;
            this.ForegroundColor = Colors.DarkRed;
            this.DisplayName = "Highlight Word 4";
            this.ZOrder = 5;
        }
    }

    [Export(typeof(EditorFormatDefinition))]
    [Name("MarkerFormatDefinition/HighlightWordFormatDefinition5")]
    [UserVisible(true)]
    internal class HighlightWordFormatDefinition5 : MarkerFormatDefinition
    {
        public HighlightWordFormatDefinition5()
        {
            this.BackgroundColor = Colors.LimeGreen;
            this.ForegroundColor = Colors.DarkRed;
            this.DisplayName = "Highlight Word 5";
            this.ZOrder = 5;
        }
    }

    #endregion

    class HighlightWordClass : TextMarkerTag
    {
        public HighlightWordClass(int color) : base($"MarkerFormatDefinition/HighlightWordFormatDefinition{(color % 5) + 1}") { }
    }

    internal class HighlightWordTagger : ITagger<HighlightWordClass>
    {
        public HighlightWordTagger(ITextView view, ITextBuffer sourceBuffer, ITextSearchService textSearchService,
ITextStructureNavigator textStructureNavigator)
        {
            this.View = view;
            this.SourceBuffer = sourceBuffer;
            this.TextSearchService = textSearchService;
            this.TextStructureNavigator = textStructureNavigator;
            this.WordSpans = new List<NormalizedSnapshotSpanCollection>();
            this.View.LayoutChanged += ViewLayoutChanged;
            HighlightWordsSettingsManager.SettingsChanged += () => { UpdateWordsAdornments(); };

            UpdateAtCaretPosition(this.View.Caret.Position);
        }

        void ViewLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            // If a new snapshot wasn't generated, then skip this layout  
            if (e.NewSnapshot != e.OldSnapshot)
            {
                UpdateAtCaretPosition(View.Caret.Position);
            }
        }

        //void CaretPositionChanged(object sender, CaretPositionChangedEventArgs e)
        //{
        //    UpdateAtCaretPosition(e.NewPosition);
        //}

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;

        void UpdateAtCaretPosition(CaretPosition caretPosition)
        {
            SnapshotPoint? point = caretPosition.Point.GetPoint(SourceBuffer, caretPosition.Affinity);

            if (!point.HasValue)
                return;

            RequestedPoint = point.Value;
            UpdateWordsAdornments();
        }



        void UpdateWordsAdornments()
        {
            SnapshotPoint currentRequest = RequestedPoint;
            List<NormalizedSnapshotSpanCollection> wordSpansList = new List<NormalizedSnapshotSpanCollection>();
            var words = HighlightWordsSettingsManager.GetWords();
            foreach (string word in words)
            {
                //Find the new spans  
                FindData findData = new FindData(word, SourceBuffer.CurrentSnapshot)
                {
                    FindOptions = FindOptions.WholeWord | FindOptions.MatchCase
                };

                wordSpansList.Add(new NormalizedSnapshotSpanCollection(TextSearchService.FindAll(findData)));
            }

            //If another change hasn't happened, do a real update   
            if (currentRequest == RequestedPoint)
                SynchronousUpdate(currentRequest, wordSpansList);
        }

        public static bool WordExtentIsValid(SnapshotPoint currentRequest, TextExtent word)
        {
            return word.IsSignificant
                && currentRequest.Snapshot.GetText(word.Span).Any(c => char.IsLetter(c));
        }

        void SynchronousUpdate(SnapshotPoint currentRequest, List<NormalizedSnapshotSpanCollection> newSpans)
        {
            lock (updateLock)
            {
                if (currentRequest != RequestedPoint)
                    return;

                WordSpans = newSpans;

                TagsChanged?.Invoke(this, new SnapshotSpanEventArgs(new SnapshotSpan(SourceBuffer.CurrentSnapshot, 0, SourceBuffer.CurrentSnapshot.Length)));
            }
        }

        public IEnumerable<ITagSpan<HighlightWordClass>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            List<NormalizedSnapshotSpanCollection> wordSpansList = WordSpans;
            if (spans.Count == 0 || wordSpansList.Count == 0)
                yield break;

            for (int i = 0; i < wordSpansList.Count; ++i)
            {
                if (wordSpansList[i].Count == 0) continue;
                // If the requested snapshot isn't the same as the one our words are on, translate our spans to the expected snapshot   
                if (spans[0].Snapshot != wordSpansList[i][0].Snapshot)
                {
                    wordSpansList[i] = new NormalizedSnapshotSpanCollection(
                        wordSpansList[i].Select(span => span.TranslateTo(spans[0].Snapshot, SpanTrackingMode.EdgeExclusive)));

                    // currentWord = currentWord.TranslateTo(spans[0].Snapshot, SpanTrackingMode.EdgeExclusive);
                }

                // Second, yield all the other words in the file   
                foreach (SnapshotSpan span in NormalizedSnapshotSpanCollection.Overlap(spans, wordSpansList[i]))
                {
                    yield return new TagSpan<HighlightWordClass>(span, new HighlightWordClass(i));
                }
            }
        }

        ITextView View { get; set; }
        ITextBuffer SourceBuffer { get; set; }
        ITextSearchService TextSearchService { get; set; }
        ITextStructureNavigator TextStructureNavigator { get; set; }
        List<NormalizedSnapshotSpanCollection> WordSpans { get; set; }
        SnapshotPoint RequestedPoint { get; set; }
        object updateLock = new object();
    }

    [Export(typeof(IViewTaggerProvider))]
    [ContentType("text")]
    [TagType(typeof(TextMarkerTag))]
    internal class HighlightWordTaggerProvider : IViewTaggerProvider
    {
        [Import]
        internal ITextSearchService TextSearchService { get; set; }

        [Import]
        internal ITextStructureNavigatorSelectorService TextStructureNavigatorSelector { get; set; }

        public ITagger<T> CreateTagger<T>(ITextView textView, ITextBuffer buffer) where T : ITag
        {
            //provide highlighting only on the top buffer   
            if (textView.TextBuffer != buffer)
                return null;

            // HighlightWordCommands.Instance.TextStructureNavigatorSelectorService = TextStructureNavigatorSelector;
            ITextStructureNavigator textStructureNavigator =
                TextStructureNavigatorSelector.GetTextStructureNavigator(buffer);
            var tagger = new HighlightWordTagger(textView, buffer, TextSearchService, textStructureNavigator);

            return tagger as ITagger<T>;
        }
    }

}