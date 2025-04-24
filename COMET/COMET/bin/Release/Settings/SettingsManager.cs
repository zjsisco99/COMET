using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xaml;
using System.Resources;
using System.Reflection;

namespace COMET.Settings
{
    internal class SettingsManager
    {
        public string settingsFullName { get; set; } = default;
        public string settingsInitials { get; set; } = default;
        public string settingsInnerTag { get; set; } = default;
        public string settingsOuterProperty { get; set; } = default;
        public string settingsOuterTag { get; set; } = default;
        public string settingsUserText { get; set; } = default;
        public string sectionCreatedBy { get; set; } = default;
        public String[] names { get; set; } =
        {
            "Created By",
            "Events Handled",
            "Event Raised",
            "Exception Thrown",
            "Exception Caught",
            "Namespace",
            "Purpose",
            "Revise History"
        };

        public String[] toggles { get; set; } =
        {
            "Y", //00 - toggleClassCB        
            "Y", //01 - toggleClassPurpose  
            "Y", //02 - toggleClassRH        
            "Y", //03 - toggleConstructCB   
            "Y", //04 - toggleConstructEvents
            "Y", //05 - toggleConstructExcepts
            "Y", //06 - toggleConstructPurpose
            "Y", //07 - toggleConstructRH   
            "Y", //08 - toggleEnumCB        
            "Y", //09 - toggleEnumPurpose    
            "Y", //10 - toggleEnumRH         
            "Y", //11 - toggleFileCB         
            "Y", //12 - toggleFilePurpose    
            "Y", //13 - toggleFileRH         
            "Y", //14 - toggleMethodCB       
            "Y", //15 - toggleMethodEvents   
            "Y", //16 - toggleMethodExcepts  
            "Y", //17 - toggleMethodPurpose  
            "Y", //18 - toggleMethodRH       
            "Y", //19 - toggleNamespaceCB    
            "Y", //20 - toggleNamespacePurpose
            "Y", //21 - toggleNamespaceRH    
            "Y", //22 - toggleOBJCB          
            "Y", //23 - toggleOBJPurpose    
            "Y"  //24 - toggleOBJRH          
        };
   
        public SettingsManager()
        {

        }

    }
}
