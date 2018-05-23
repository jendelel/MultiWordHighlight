using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MultiWordHighlight
{
    internal static class HighlightWordsSettingsManager
    {
        // Because my code is always called from the UI thred, this succeeds.  
        internal static SettingsManager VsManagedSettingsManager =
            new ShellSettingsManager(ServiceProvider.GlobalProvider);

        private const int _maxWords = 5;
        private const string _collectionSettingsName = "Text Editor";
        private const string _settingName = "HighlightWords";

        static internal bool ToggleWord(string word)
        {
            if (!IsValidWord(word))
                throw new ArgumentOutOfRangeException(
                    "column", "The string must contain only one nonempty word.");
            var words = HighlightWordsSettingsManager.GetWords();
            if (words.Contains(word))
            {
                // Remove the word.
                words.Remove(word);
            }
            else
            {
                if (words.Count() >= _maxWords) return false;
                words.Add(word);
            }
            WriteSettings(words);
            return true;
        }

        static internal bool CanToggleWord(string word)
        {
            if (!IsValidWord(word))
                return false;
            var words = GetWords();
            if (words.Contains(word))
            {
                return true;
            }
            else
            {
                return words.Count < _maxWords;
            }
        }

        static internal void RemoveAllWords()
        {
            WriteSettings(new List<string>());
        }

        private static bool IsValidWord(string word)
        {
            // zero is allowed (per user request)  
            return word != null && word.Length > 0 && word.Trim() == word && word.IndexOf(' ') < 0;
        }

        static private string _highlightWordsConfiguration;
        static private string HighlightWordsConfiguration
        {
            get
            {
                if (_highlightWordsConfiguration == null)
                {
                    _highlightWordsConfiguration =
                        GetUserSettingsString(
                            HighlightWordsSettingsManager._collectionSettingsName,
                            HighlightWordsSettingsManager._settingName)
                        .Trim();
                }
                return _highlightWordsConfiguration;
            }

            set
            {
                if (value != _highlightWordsConfiguration)
                {
                    _highlightWordsConfiguration = value;
                    WriteUserSettingsString(
                        HighlightWordsSettingsManager._collectionSettingsName,
                        HighlightWordsSettingsManager._settingName, value);
                    // Notify ColumnGuideAdornments to update adornments in views.  
                    HighlightWordsSettingsManager.SettingsChanged?.Invoke();
                }
            }
        }

        internal static string GetUserSettingsString(string collection, string setting)
        {
            var store = HighlightWordsSettingsManager
                            .VsManagedSettingsManager
                            .GetReadOnlySettingsStore(SettingsScope.UserSettings);
            return store.GetString(collection, setting, "");
        }

        internal static void WriteUserSettingsString(string key, string propertyName,
                                                     string value)
        {
            var store = HighlightWordsSettingsManager
                            .VsManagedSettingsManager
                            .GetWritableSettingsStore(SettingsScope.UserSettings);
            store.CreateCollection(key);
            store.SetString(key, propertyName, value);
        }

        // Persists settings and sets property with side effect of signaling  
        // ColumnGuideAdornments to update.  
        static private void WriteSettings(List<string> words)
        {
            string value = ComposeSettingsString(words);
            HighlightWordsConfiguration = value;
        }

        private static string ComposeSettingsString(List<string> words)
        {
            var ser = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(List<string>));
            using (MemoryStream memstream = new MemoryStream())
            {
                ser.WriteObject(memstream, words);
                memstream.Position = 0;
                return (new StreamReader(memstream)).ReadToEnd();
            }
        }

        static internal List<string> GetWords()
        {
            string settings = HighlightWordsSettingsManager.HighlightWordsConfiguration;
            if (String.IsNullOrEmpty(settings))
                return new List<string>();
            List<string> result = new List<string>();
            using (var memStream = new MemoryStream())
            {
                var streamWriter = new StreamWriter(memStream);
                streamWriter.Write(settings);
                streamWriter.Flush();
                memStream.Position = 0;
                string toBeSure = (new StreamReader(memStream)).ReadToEnd();
                memStream.Position = 0;
                var ser = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(List<string>));
                result = (List<string>)ser.ReadObject(memStream);
            }
            return result;
        }

        // Delegate and Event to fire when settings change so that ColumnGuideAdornments   
        // can update.  We need nothing special in this event since the settings manager   
        // is statically available.  
        //  
        internal delegate void SettingsChangedHandler();
        static internal event SettingsChangedHandler SettingsChanged;
    }
}