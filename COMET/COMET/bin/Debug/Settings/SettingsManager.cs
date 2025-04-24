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

namespace COMET.Settings
{
    internal class SettingsManager
    {
        static readonly string APP_DATA_DIRECTORY = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "COMET", "Settings");
        static readonly string SETTINGS_FILE = Path.Combine(APP_DATA_DIRECTORY, "settings.json");

        public string settingsFullName { get; set; } = default;
        public string settingsInitials { get; set; } = default;

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
        
        public Dictionary<string, string> colors { get; set; } = new Dictionary<string, string>
        {
                {"InnerTag", "#FFB2CAB2"},
                {"OuterTag", "#FF008000"},
                {"OuterProperty", "#FF03C203"},
                {"UserText", "#ffffff"},
        };

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

        public SettingsManager()
        {
            CheckDirectory(APP_DATA_DIRECTORY);
        }

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

        /// <summary>
        /// Description:
        /// Reads the Settings JSON file and loads the settings into the GUI
        /// ==========================================
        /// Called By: 
        /// ToolWindow1Control
        /// ==========================================
        /// Revision History:
        /// date        who     description
        /// ================================
        /// 2/9/2025   zs      inital creation
        /// </summary>
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

        /// <summary>
        /// Description:
        /// Saves the Settings OBJ to the Settings JSON file
        /// ==========================================
        /// Called By: 
        /// ==========================================
        /// Revision History:
        /// date        who     description
        /// ================================
        /// 2/9/2025   zs      inital creation
        /// </summary>
        public void SaveSettings()
        {
            //write JSON
            var options = new JsonSerializerOptions();
            options.WriteIndented = true;
            string jsonString = JsonSerializer.Serialize(this, options);
            string fullPath = SETTINGS_FILE;
            File.WriteAllText(fullPath, jsonString);
        }

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

        public string getFullName()
        {
            return this.settingsFullName;
        }

        public void setFullName(string fullName)
        {
            this.settingsFullName = fullName;
        }
        
        public string getInitials()
        {
            return this.settingsInitials;
        }

        public void setInitials(string initials)
        {
            this.settingsInitials = initials;
        }

        public Color getColor(string color)
        {
            return (Color)System.Windows.Media.ColorConverter.ConvertFromString(this.colors[color]);
        }

        public void setColor(string option, string color)
        {
            this.colors[option] = color;
        }

        public string getName(string option)
        {
            return this.names[option];
        }

        public void setName(string option, string name)
        {
            this.names[option] = name;
        }

        public bool getToggle(string option)
        {
            return this.toggles[option];
        }

        public void setToggle(string option, bool toggle)
        {
            this.toggles[option] = toggle;
        }
    }
}
