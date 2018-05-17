// DataBuffer is an ArcGIS add-in used to create 'species alert'
// layers from existing species data.
//
// Copyright © 2017 SxBRC, 2017-2018 TVERC
//
// This file is part of DataBuffer.
//
// DataBuffer is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataBuffer is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataBuffer.  If not, see <http://www.gnu.org/licenses/>.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;

namespace HLStringFunctions
{
    /// Original code from:
    /// http://www.codeproject.com/Articles/11556/Converting-Wildcards-to-Regexes<summary>
    /// Represents a wildcard running on the
    /// <see cref="System.Text.RegularExpressions"/> engine.
    /// </summary>
    public class Wildcard : Regex
    {
        /// <summary>
        /// Initializes a wildcard with the given search pattern.
        /// </summary>
        /// <param name="pattern">The wildcard pattern to match.</param>
        public Wildcard(string pattern)
            : base(WildcardToRegex(pattern))
        {
        }

        /// <summary>
        /// Initializes a wildcard with the given search pattern and options.
        /// </summary>
        /// <param name="pattern">The wildcard pattern to match.</param>
        /// <param name="options">A combination of one or more
        /// <see cref="System.Text.RegexOptions"/>.</param>
        public Wildcard(string pattern, RegexOptions options)
            : base(WildcardToRegex(pattern), options)
        {
        }

        /// <summary>
        /// Converts a wildcard to a regex.
        /// </summary>
        /// <param name="pattern">The wildcard pattern to convert.</param>
        /// <returns>A regex equivalent of the given wildcard.</returns>
        public static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern).
             Replace("\\*", ".*").
             Replace("\\?", ".") + "$";
        }
    }

    public class StringFunctions
    {
        public StringFunctions()
        {
            // constructor takes no arguments.
        }

        public string StripIllegals(string anInputString, string aReplaceString, bool isFileName=false)
        {
            // Function to remove all potentially special characters from a string and return the result.

            // If it is a file name, check if there is a '.' at fourth place before last.
            bool blAddFileDot = false;
            if (isFileName)
            {
                char chTest = anInputString[anInputString.Length - 4];
                if (chTest == '.') blAddFileDot = true;
            }

            string strOutputString = anInputString;
            List<string> theIllegals = new List<string>() { @"\", "%", "$", ":", "*", "/", "?", "<", ">", "|", "~", "£" , "."};
            foreach (string aSearchString in theIllegals)
            {
                strOutputString = strOutputString.Replace(aSearchString, aReplaceString);
            }
            if (blAddFileDot)
            {
                strOutputString  = strOutputString.Remove(strOutputString.Length - 4, 1);
                strOutputString = strOutputString.Insert(strOutputString.Length - 3, ".");
            }
            return strOutputString;
        }

        public bool isValid(string aReplacementCharacter)
        {
            List<string> theIllegals = new List<string>() { @"\", "%", "$", ":", "*", "/", "?", "<", ">", "|", "~", "£", "." };

            if (theIllegals.IndexOf(aReplacementCharacter) == -1)
                return true;
            return false;

        }
        public string KeepNumbersAndSpaces(string anInputString, string aReplaceCharacter)
        {
            // Does what it says on the packet - keeps numbers and spaces from an input string and returns the results.
            string strOutputString = "";
            int a;
            int aCount = 0;
            foreach (char strTest in anInputString)
            {
                if (int.TryParse(strTest.ToString(), out a) == true)
                {
                    strOutputString += strTest.ToString();
                    aCount = aCount + 1;
                }
                // Replace characters and spaces are not included at the start of the reference.
                else if ((strTest == ' ' || strTest.ToString() == aReplaceCharacter) && aCount > 0)
                    strOutputString += strTest.ToString();
            }
            return strOutputString;
        }

        public string GetSubref(string anInputString, string aReplaceString)
        {
            // Gets the sub-reference out of a short reference string. 
            // Input should look like xx.xxxx or xxxx where x is an integer.
            int a = anInputString.IndexOf(aReplaceString) + 1; // The index of the first numeric character after the replace character.
            return anInputString.Substring(a, anInputString.Length - a); 
        }

        public string getGroupColumnsFormatted(string aGroupColumnString)
        {
            List<string> strColumns = aGroupColumnString.Split(',').ToList();
            string strFormatted = "";
            foreach (string strEntry in strColumns)
            {
                strFormatted = strFormatted + strEntry.Trim() + ";";
            }
            if (strFormatted != "")
                return strFormatted.Substring(0, strFormatted.Length - 1); // Remove the final semicolon.
            else
                return "";

        }

        public string getStatsColumnsFormatted(string aStatsColumnString)
        {
            List<string> strEntries = aStatsColumnString.Split('$').ToList();
            string strFormatted = "";
            foreach (string strEntry in strEntries)
            {
                strFormatted = strFormatted + strEntry.Replace(";", " ") + ";";
            }
            if (strFormatted != "")
                return strFormatted.Substring(0, strFormatted.Length - 1); // Remove the final comma.
            else
                return "";
        }

        public string AlignStatsColumns(string AllColumns, string StatsColumns, string GroupColumns)
        {
            if (GroupColumns == "" || AllColumns == "") //GroupColumns == "" || 
                return StatsColumns;
            List<string> liAllColumns = AllColumns.Split(',').ToList();
            foreach (string strFieldName in liAllColumns)
            {
                string strFieldNameTr = strFieldName.Trim();
                if (strFieldNameTr.Substring(0, 1) != "\"")
                {
                    // Is it in the group columns?
                    if (GroupColumns.IndexOf(strFieldNameTr, StringComparison.CurrentCultureIgnoreCase) == -1)
                    {
                        // Is it in the stats columns?
                        if (StatsColumns.IndexOf(strFieldNameTr, StringComparison.CurrentCultureIgnoreCase) == -1)
                        {
                            // It is in neither - add it.
                            if (StatsColumns != "")
                                StatsColumns = StatsColumns + ";";
                            StatsColumns = StatsColumns + strFieldNameTr + " FIRST";
                        }
                    }
                }
            }

            return StatsColumns;
        }

        public List<string> ExtractGroups(List<string> LayerList)
        {
            // This function looks at each layer in LayerList and extracts unique group names (in front of any hyphen in the name) 
            List<string> liGroups = new List<string>();
            foreach (string strLayerName in LayerList)
            {
                int intHyphenIndex = strLayerName.IndexOf('-');
                if (intHyphenIndex != -1) // It has a group name
                {
                    string strGroupName = strLayerName.Substring(0, intHyphenIndex); // Check if we already have this one
                    if (liGroups.IndexOf(strGroupName) == -1)
                    {
                        liGroups.Add(strGroupName); // If not, add it.
                    }
                }
            }
            return liGroups; // Return the list of group names.

        }

        public string GetGroupName(string LayerName)
        {
            int intHyphenIndex = LayerName.IndexOf('-');
            if (intHyphenIndex != -1) // It has a group name
            {
                string strGroupName = LayerName.Substring(0, intHyphenIndex);
                return strGroupName; // Return the group name
            }
            else
                return ""; // No group name.
        }

        public string ReplaceSearchStrings(string RawName, string Reference, string SiteName, string ShortRef, string Subref)
        {
            string CleanName = RawName;
            CleanName = CleanName.Replace("%ref%", Reference);
            CleanName = CleanName.Replace("%shortref%", ShortRef);
            CleanName = CleanName.Replace("%subref%", Subref);
            CleanName = CleanName.Replace("%sitename%", SiteName);
            // Take account of the occurrence of dangling underscores (if no site name was given).
            if (CleanName.Substring(CleanName.Length - 1, 1) == "_")
                CleanName = CleanName.Substring(0, CleanName.Length - 1);
            return CleanName;
        }
    }


}


