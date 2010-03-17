using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml;

namespace NotesToGoogle
{
    class SyncPreferences
    {
        /// <summary>
        /// Class constructor for SyncPreferences class.
        /// This class is used to save and retrieve the user preferences from file.
        /// </summary>
        public SyncPreferences()
        {
            // Initialize the storage data type(s)
            htSyncPreferences = new Hashtable();
        }

        /// <summary>
        /// Method used to save / write preferences to file
        /// </summary>
        public void SavePreferences ()
        {
            try
            {
                // Check to see if the file exists, if so delete it
                if (File.Exists(sPrefFile))
                {
                    // Delete the file if one exists
                    File.Delete(sPrefFile);
                }

                // Initialize XmlTextWriter to output file
                XmlTextWriter xPrefWriter = new XmlTextWriter(sPrefFile, null);
                xPrefWriter.Formatting = Formatting.Indented;

                // Begin writing file
                xPrefWriter.WriteStartDocument();
                xPrefWriter.WriteComment(" NotesToGoogleCal Preferences File ");
                xPrefWriter.WriteComment(" Date:                             ");
                xPrefWriter.WriteStartElement("Settings");
                xPrefWriter.WriteWhitespace("\n");

                foreach (DictionaryEntry de in htSyncPreferences)
                {
                    //xPrefWriter.WriteElementString(de.Key.ToString(), de.Value.ToString());
                    //xPrefWriter.WriteWhitespace("\n");

                    String _element = "\t<Setting name=\"" + de.Key.ToString() + "\">" + de.Value.ToString() + "</Setting>\n";
                    xPrefWriter.WriteRaw(_element);
                }

                xPrefWriter.WriteEndElement(); // </Preferences>
                xPrefWriter.WriteEndDocument();
                xPrefWriter.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Error", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// Method used to load / read preferences from file
        /// </summary>
        public Boolean LoadPreferences()
        {
            try
            {
                if (File.Exists(sPrefFile))
                {
                    // Create the reader
                    XmlTextReader xPrefReader = new XmlTextReader(sPrefFile);

                    //Loop through the file reading in Element by element
                    while (xPrefReader.Read())
                    {
                        if ((xPrefReader.IsStartElement()) && (xPrefReader.Name == "Setting"))
                        {
                            // Insert into the Hash table with Key=Attribute and Value=ElementAsString()
                            htSyncPreferences.Add(xPrefReader.GetAttribute(0), xPrefReader.ReadElementContentAsString());
                        }
                    }

                    xPrefReader.Close();
                    return true;
                }

                return false;
            } 
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK);
                return false;
            }
        }

        /// <summary>
        /// Method used to set preference values
        /// </summary>
        /// <param name="_prefName">Preference Name to store</param>
        /// <param name="_prefValue">Preference Value to store</param>
        public void SetPreference(String _prefName, String _prefValue)
        {
            // Check to see if preference is already set
            if (htSyncPreferences.ContainsKey(_prefName))
            {
                // If the preference exists we remove the old key/value
                htSyncPreferences.Remove(_prefName);
            }

            // Add the preference
            htSyncPreferences.Add(_prefName, _prefValue);
        }

        /// <summary>
        /// Method used to retrieve preference value
        /// </summary>
        /// <param name="_prefName">Preference name to return value of</param>
        /// <returns>String value of preference</returns>
        public String GetPreference(String _prefName)
        {
            String _prefValue;

            if (htSyncPreferences.Contains(_prefName))
            {
                _prefValue = htSyncPreferences[_prefName].ToString();
            } else {
                _prefValue = "";
            }

            return _prefValue;
        }

        // Class variables
        Hashtable htSyncPreferences;
        String sPrefFile = "NotesToGoogleCal.preference";
    }
}
