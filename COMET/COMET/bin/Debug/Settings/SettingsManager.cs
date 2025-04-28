using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xaml;
using System.Resources;
using System.Reflection;
using System.IO;
using System.Runtime;
using System.Text.Json;
using System.Drawing;
using System.Windows.Media;
using Color = System.Windows.Media.Color;
using System.Diagnostics;
/// <FLOWERBOX file="SettingsManager.cs">
/// <Created_By>
/// COMET DEV TEAM
/// </Created_By>
/// <Purpose>
/// Controls the settings for the COMET extension.
/// </Purpose>
/// <Revise_History>
/// 4/27/2025 - Initial release
/// </Revise_History>
/// </FLOWERBOX>

/// <NAMESPACE name="COMET.Settings">
/// <Purpose>
/// Namespace for the COMET extension settings.
/// </Purpose>
/// </NAMESPACE>
namespace COMET.Settings
{
    /// <CLASS name="SettingsManager">
    /// <Purpose>
    /// Creates instance of the settings manager.
    /// </Purpose>
    /// </CLASS>
    internal class SettingsManager
    {
        static readonly string APP_DATA_DIRECTORY = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "COMET", "Settings");
        static readonly string SETTINGS_FILE = Path.Combine(APP_DATA_DIRECTORY, "settings.json");

       /// <PROPERTY name="settingsFullName">
       /// <Purpose>
       /// UNUSED
       /// </Purpose>
       /// </PROPERTY>
        public string settingsFullName { get; set; } = default;

       /// <PROPERTY name="settingsInitials">
       /// <Purpose>
       /// USERNAME
       /// </Purpose>
       /// </PROPERTY>
        public string settingsInitials { get; set; } = default;

       /// <PROPERTY name="names">
       /// <Purpose>
       /// TAG NAMES
       /// </Purpose>
       /// </PROPERTY>
        public Dictionary<string, string> names { get; set; } = new Dictionary<string, string>
            {
                {"Created By", "Created_By"},
                {"Events Handled", "Events_Handled"},
                {"Event Raised", "Event_Raised"},
                {"Exception Thrown", "Exception_Thrown"},
                {"Exception Caught", "Exception_Caught"},
                {"Purpose", "Purpose"},
                {"Revise History", "Revise_History"}
            };
        
       /// <PROPERTY name="colors">
       /// <Purpose>
       /// TAG COLORS
       /// </Purpose>
       /// </PROPERTY>
        public Dictionary<string, string> colors { get; set; } = new Dictionary<string, string>
        {
                {"InnerTag", "#FFB2CAB2"},
                {"OuterTag", "#FF008000"},
                {"OuterProperty", "#FF03C203"},
                {"UserText", "#ffffff"},
        };

        /// <PROPERTY name="toggles">
        /// <Purpose>
        /// TAG TOGGLES
        /// </Purpose>
        /// </PROPERTY>
        public Dictionary<string, bool> toggles { get; set; } = new Dictionary<string, bool>
        {
            {"toggleNamespaceCB", true },
            {"toggleNamespacePurpose", true },
            {"toggleNamespaceRH", true },
            {"toggleFileCB", true },
            {"toggleFilePurpose", true },
            {"toggleFileRH", true },
            {"toggleEnumCB", true },
            {"toggleEnumPurpose", true },
            {"toggleEnumRH", true },
            {"toggleOBJCB", true },
            {"toggleOBJPurpose", true },
            {"toggleOBJRH", true },
            {"toggleClassCB", true },
            {"toggleClassPurpose", true },
            {"toggleClassRH", true },
            {"toggleMethodCB", true },
            {"toggleMethodEvents", true },
            {"toggleMethodExcepts", true },
            {"toggleMethodPurpose", true },
            {"toggleMethodRH", true },
            {"toggleStructCB", true },
            {"toggleStructEvents", true },
            {"toggleStructExcepts", true },
            {"toggleStructPurpose", true },
            {"toggleStructRH", true }
        };

        /// <METHOD name="SettingsManager">
        /// <Purpose> 
        /// Creates instance of the settings manager.
        /// </Purpose>
        /// <Parameters> 
        /// 
        /// </Parameters>
        /// </METHOD>
        public SettingsManager()
        {
            CheckDirectory(APP_DATA_DIRECTORY);
        }

        /// <METHOD name="CheckDirectory">
        /// <Purpose> 
        /// Checks if the AppData directory exists. If not, creates it.
        /// </Purpose>
        /// <Parameters>
        ///    directory(string):
        /// </Parameters>
        /// </METHOD>
        public void CheckDirectory(string directory)
        {
            DirectoryInfo di = new DirectoryInfo(directory);
            if (!di.Exists)
            {
                di.Create();
                SaveSettings();
                LoadSettings();
            }
            else if (di.Exists)
            {
                LoadSettings();
            }
        }

        /// <METHOD name="LoadSettings">
        /// <Purpose> 
        /// Loads the settings from the AppData directory.
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public void LoadSettings()
        {
            //read JSON
            string fullPath = SETTINGS_FILE;
            string readJson = File.ReadAllText(fullPath);
            var options = new { settingsFullName = "",
                                settingsInitials = "",
                                names = new Dictionary<string, string>(),
                                colors = new Dictionary<string, string>(),
                                toggles = new Dictionary<string, bool>()
                              };
            var contents = JsonSerializer.Deserialize(readJson, options.GetType());
            this.settingsFullName = (string)contents.GetType().GetProperty("settingsFullName").GetValue(contents);
            this.settingsInitials = (string)contents.GetType().GetProperty("settingsInitials").GetValue(contents);
            this.names = (Dictionary<string, string>)contents.GetType().GetProperty("names").GetValue(contents);
            this.colors = (Dictionary<string, string>)contents.GetType().GetProperty("colors").GetValue(contents);
            this.toggles = (Dictionary<string, bool>)contents.GetType().GetProperty("toggles").GetValue(contents);
        }

       /// <METHOD name="SaveSettings">
       /// <Purpose> 
       /// Saves the settings in the AppData directory
       /// </Purpose>
       /// <Parameters> 
       ///
       /// </Parameters>
       /// </METHOD>
        public void SaveSettings()
        {
            //write JSON
            var options = new JsonSerializerOptions();
            options.WriteIndented = true;
            string jsonString = JsonSerializer.Serialize(this, options);
            string fullPath = SETTINGS_FILE;
            File.WriteAllText(fullPath, jsonString);
        }

        /// <METHOD name="DefaultSettings">
        /// <Purpose> 
        ///  Sets the default settings for the COMET extension.
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public void DefaultSettings()
        {
            this.setFullName(string.Empty);
            this.setInitials(string.Empty);
            this.setColor("InnerTag", "#FFB2CAB2");
            this.setColor("OuterTag", "#FF008000");
            this.setColor("OuterProperty", "#FF03C203");
            this.setColor("UserText", "#ffffff");
            this.setName("Created By", "Created_By");
            this.setName("Events Handled", "Events_Handled");
            this.setName("Event Raised", "Event_Raised");
            this.setName("Exception Thrown", "Exception_Thrown");
            this.setName("Exception Caught", "Exception_Caught");
            this.setName("Purpose", "Purpose");
            this.setName("Revise History", "Revise_History");
            this.setToggle("toggleNamespaceCB", true);
            this.setToggle("toggleClassCB", true);
            this.setToggle("toggleClassPurpose", true);
            this.setToggle("toggleClassRH", true);
            this.setToggle("toggleStructCB", true);
            this.setToggle("toggleStructEvents", true);
            this.setToggle("toggleStructExcepts", true);
            this.setToggle("toggleStructPurpose", true);
            this.setToggle("toggleStructRH", true);
            this.setToggle("toggleEnumCB", true);
            this.setToggle("toggleEnumPurpose", true);
            this.setToggle("toggleEnumRH", true);
            this.setToggle("toggleFileCB", true);
            this.setToggle("toggleFilePurpose", true);
            this.setToggle("toggleFileRH", true);
            this.setToggle("toggleMethodCB", true);
            this.setToggle("toggleMethodEvents", true);
            this.setToggle("toggleMethodExcepts", true);
            this.setToggle("toggleMethodPurpose", true);
            this.setToggle("toggleMethodRH", true);
            this.setToggle("toggleNamespaceCB", true);
            this.setToggle("toggleNamespacePurpose", true);
            this.setToggle("toggleNamespaceRH", true);
            this.setToggle("toggleOBJCB", true);
            this.setToggle("toggleOBJPurpose", true);
            this.setToggle("toggleOBJRH", true);
            this.SaveSettings();
        }

       /// <METHOD name="getFullName">
       /// <Purpose> 
       ///  UNUSED  
       /// </Purpose>
       /// <Parameters> 
       ///
       /// </Parameters>
       /// </METHOD>
        public string getFullName()
        {
            return this.settingsFullName;
        }

        /// <METHOD name="setFullName">
        /// <Purpose> 
        ///  UNUSED
        /// </Purpose>
        /// <Parameters>
        ///    fullName(string):
        /// </Parameters>
        /// </METHOD>
        public void setFullName(string fullName)
        {
            this.settingsFullName = fullName;
        }

        /// <METHOD name="getInitials">
        /// <Purpose> 
        ///  Get Username
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public string getInitials()
        {
            return this.settingsInitials;
        }

       /// <METHOD name="setInitials">
       /// <Purpose> 
       ///  Set Username
       /// </Purpose>
       /// <Parameters>
       ///    initials(string):
       /// </Parameters>
       /// </METHOD>
        public void setInitials(string initials)
        {
            this.settingsInitials = initials;
        }

        /// <METHOD name="getColor">
        /// <Purpose> 
        ///  Get Color dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    color(string):
        /// </Parameters>
        /// </METHOD>
        public Color getColor(string color)
        {
            return (Color)System.Windows.Media.ColorConverter.ConvertFromString(this.colors[color]);
        }

        /// <METHOD name="setColor">
        /// <Purpose> 
        ///  Set color dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    option(string):
        ///    color(string):
        /// </Parameters>
        /// </METHOD>
        public void setColor(string option, string color)
        {
            this.colors[option] = color;
        }

        /// <METHOD name="getName">
        /// <Purpose> 
        ///  Get name dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    option(string):
        /// </Parameters>
        /// </METHOD>
        public string getName(string option)
        {
            return this.names[option];
        }

        /// <METHOD name="setName">
        /// <Purpose> 
        ///  set name dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    option(string):
        ///    name(string):
        /// </Parameters>
        /// </METHOD>
        public void setName(string option, string name)
        {
            this.names[option] = name;
        }

        /// <METHOD name="getToggle">
        /// <Purpose> 
        ///  get toggle dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    option(string):
        /// </Parameters>
        /// </METHOD>
        public bool getToggle(string option)
        {
            return this.toggles[option];
        }

        /// <METHOD name="setToggle">
        /// <Purpose> 
        ///  set toggle dictionary option
        /// </Purpose>
        /// <Parameters>
        ///    option(string):
        ///    toggle(bool):
        /// </Parameters>
        /// </METHOD>
        public void setToggle(string option, bool toggle)
        {
            this.toggles[option] = toggle;
        }
    }
}
