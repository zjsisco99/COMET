using System.Windows;
using System.Windows.Controls;
using COMET.Settings;
using System.IO;
using System.Text.Json;
using System;
using System.Diagnostics;
using Microsoft.VisualStudio.PlatformUI;
using Syncfusion.Windows.Shared;
using System.IO.Packaging;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System;
using EnvDTE80;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Linq;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Media;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Package = Microsoft.VisualStudio.Shell.Package;
using System.Drawing;
using Color = System.Windows.Media.Color;
using Brush = System.Windows.Media.Brush;
using System.Collections.Generic;
using COMET.Resources;
using Newtonsoft.Json.Linq;
using System.Reflection;

/// <FLOWERBOX file="ToolWindow1Control.xaml.cs">
/// <Created_By>
/// COMET DEV TEAM
/// </Created_By>
/// <Purpose>
/// Controls the backend logic of the GUI controls, and the comment generation logic.
/// </Purpose>
/// <Revise_History>
/// 4/27/2025 - Initial release
/// </Revise_History>
/// </FLOWERBOX>

/// <NAMESPACE name="COMET">
/// <Purpose>
/// Namespace of COMET
/// </Purpose>
/// </NAMESPACE>
namespace COMET
{

      /// <CLASS name="ToolWindow1Control">
   /// <Purpose>
   ///  Class for the Tool Window
   /// </Purpose>
   /// </CLASS>
    public partial class ToolWindow1Control : UserControl
    {
        SettingsManager settings = new SettingsManager();
        XML_Tag_Names tagNames = new XML_Tag_Names();

              /// <CLASS name="Parameter">
       /// <Purpose>
       ///  Parameter class
       /// </Purpose>
       /// </CLASS>
        public class Parameter {

            /// <PROPERTY name="Name">
            /// <Purpose>
            ///  Name of Parameter
            /// </Purpose>
            /// </PROPERTY>
            public string Name { get; set; }

            /// <PROPERTY name="Type">
            /// <Purpose>
            ///  Type of Parameter
            /// </Purpose>
            /// </PROPERTY>
            public string Type { get; set; }

            /// <PROPERTY name="Description">
            /// <Purpose>
            ///  Description of Parameter
            /// </Purpose>
            /// </PROPERTY>
            public List<string> Description { get; set; } = new List<string>();
        }

        /// <STRUCTURE name="XML_Tag_Names">
        /// <Purpose>
        ///  Options for XML Tag Names
        /// </Purpose>
        /// </STRUCTURE>
        struct XML_Tag_Names
        {
            public bool initalized;
            public string tagPurpose;
            public string tagRevisionHistory;
            public string tagCreatedBy;
            public string tagEventsHandled;
            public string tagEventRaised;
            public string tagExceptionCaught;
            public string tagExceptionThrown;
        }

              /// <METHOD name="GetThemeColor">
       /// <Purpose> 
       ///  Pulls the theme that Visual Studio is using
       /// </Purpose>
       /// <Parameters> 
       ///     colorIndex(__VSSYSCOLOREX):
       /// </Parameters>
       /// <Exception_Thrown> 
       ///     InvalidOperationException:
       /// </Exception_Thrown>
       /// </METHOD>
        public static string GetThemeColor(__VSSYSCOLOREX colorIndex)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Get the IVsUIShell service
            IVsUIShell2 uiShell = (IVsUIShell2)(ServiceProvider.GlobalProvider.GetService(typeof(SVsUIShell)) as IVsUIShell);
            if (uiShell == null)
            {
                throw new InvalidOperationException("Could not obtain IVsUIShell service.");
            }

            // Get the color
            int hr = uiShell.GetVSSysColorEx((int)colorIndex, out uint color);
            string finalColor = $"#{color:X}";
            return finalColor;
        }

              /// <METHOD name="setTextColor">
       /// <Purpose> 
       ///  Sets the text color of all labels in the GUI
       /// </Purpose>
       /// <Parameters> 
       ///
       /// </Parameters>
       /// </METHOD>
        public void setTextColor()
        {
            string textColor = GetThemeColor(__VSSYSCOLOREX.VSCOLOR_COMMANDBAR_TEXT_ACTIVE);
            BrushConverter bc = new BrushConverter();
            Brush b = (Brush)bc.ConvertFromString(textColor);
            lblUpdateSelected.Foreground = b;
            lblUpdateSelected.Foreground = b;
            chkFile.Foreground = b;
            chkNamespace.Foreground = b;
            chkClass.Foreground = b;
            chkOBJProperties.Foreground = b;
            chkConstruct.Foreground = b;
            chkEnum.Foreground = b;
            chkMethod.Foreground = b;
            lblPurpose.Foreground = b;
            lblRevisionHistory.Foreground = b;
            lblEventHandled.Foreground = b;
            lblEventRaised.Foreground = b;
            lblExceptionThrown.Foreground = b;
            lblExceptionCaught.Foreground = b;
            lblCreatedBy.Foreground = b;
            lblInitials.Foreground = b;
            lblNamespace.Foreground = b;
            chkNamespacePurpose.Foreground = b;
            chkNamespaceRH.Foreground = b;
            chkNamespaceCreated.Foreground = b;
            lblFile.Foreground = b;
            chkFilePurpose.Foreground = b;
            chkFileRH.Foreground = b;
            chkFileCreated.Foreground = b;
            lblEnum.Foreground = b;
            chkEnumPurpose.Foreground = b;
            chkEnumRH.Foreground = b;
            chkEnumCreated.Foreground = b;
            lblObj.Foreground = b;
            chkObjPurpose.Foreground = b;
            chkObjRH.Foreground = b;
            chkObjCreated.Foreground = b;
            lblClass.Foreground = b;
            chkClassPurpose.Foreground = b;
            chkClassRH.Foreground = b;
            chkClassCreated.Foreground = b;
            lblMethod.Foreground = b;
            chkMethodPurpose.Foreground = b;
            chkMethodRH.Foreground = b;
            chkMethodCreated.Foreground = b;
            chkMethodEvents.Foreground = b;
            chkMethodExceptions.Foreground = b;
            lblConst.Foreground = b;
            chkConstPurpose.Foreground = b;
            chkConstRH.Foreground = b;
            chkConstCreated.Foreground = b;
            chkConstEvents.Foreground = b;
            chkConstExceptions.Foreground = b;
            lblOuterTagColor.Foreground = b;
            lblInnerTagColor.Foreground = b;
            lblOuterPropertyColor.Foreground = b;
        }

        /// <METHOD name="SetXMLColors">
        /// <Purpose> 
        ///  Sets the XML colors for the tags
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public async void SetXMLColors()
        {
            var dte = getDTE();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            Properties fontsAndColorsProperties = dte.Properties["FontsAndColors", "TextEditor"];
            FontsAndColorsItems fontsAndColorsItems = (FontsAndColorsItems)fontsAndColorsProperties.Item("FontsAndColorsItems").Object;
            string xmlDocCommentName = "XML Doc Comment - Name";
            string xmlDocCommentAttribute = "XML Doc Comment - Attribute Name";
            string xmlDocCommentComment = "XML Doc Comment - Text";
            var userColor = colorPickerOuter.Color;
            var drawingColor = System.Drawing.Color.FromArgb(userColor.A, userColor.R, userColor.G, userColor.B);
            fontsAndColorsItems.Item(xmlDocCommentName).Foreground = (uint)ColorTranslator.ToWin32(drawingColor);
            userColor = colorPickerProperty.Color;
            drawingColor = System.Drawing.Color.FromArgb(userColor.A, userColor.R, userColor.G, userColor.B);
            fontsAndColorsItems.Item(xmlDocCommentAttribute).Foreground = (uint)ColorTranslator.ToWin32(drawingColor);
            userColor = colorPickerInner.Color;
            drawingColor = System.Drawing.Color.FromArgb(userColor.A, userColor.R, userColor.G, userColor.B);
            fontsAndColorsItems.Item(xmlDocCommentComment).Foreground = (uint)ColorTranslator.ToWin32(drawingColor);
        }

        /// <METHOD name="ToolWindow1Control">
        /// <Purpose> 
        ///  Initializes the GUI
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public ToolWindow1Control()
        {
            this.InitializeComponent();
            InitialGridVisibility();
            setTextColor();
            SetXMLColors();
        }

       /// <METHOD name="LoadSettings">
       /// <Purpose> 
       ///  Loads the settings from the AppData folder
       /// </Purpose>
       /// <Parameters> 
       ///
       /// </Parameters>
       /// </METHOD>
        public void LoadSettings()
        {
            settings.LoadSettings();
            //Settings screen
            txtInitials.Text = settings.getInitials();
            colorPickerProperty.Color = settings.getColor("OuterProperty");
            colorPickerOuter.Color = settings.getColor("OuterTag");
            colorPickerInner.Color = settings.getColor("InnerTag");


            //SectionName screen
            txtPurpose.Text = settings.getName("Purpose");
            txtRevisionHistory.Text = settings.getName("Revise History");
            txtEventHandled.Text = settings.getName("Events Handled");
            txtEventRaised.Text = settings.getName("Event Raised");
            txtExceptionCaught.Text = settings.getName("Exception Caught");
            txtExceptionThrown.Text = settings.getName("Exception Thrown");
            txtCreatedBy.Text = settings.getName("Created By");

            //SectionToggle screen
            ValidateToggleCheckboxes(chkNamespacePurpose, settings.getToggle("toggleNamespacePurpose"));
            ValidateToggleCheckboxes(chkNamespaceRH, settings.getToggle("toggleNamespaceRH"));
            ValidateToggleCheckboxes(chkNamespaceCreated, settings.getToggle("toggleNamespaceCB"));
            ValidateToggleCheckboxes(chkFilePurpose, settings.getToggle("toggleFilePurpose"));
            ValidateToggleCheckboxes(chkFileRH, settings.getToggle("toggleFileRH"));
            ValidateToggleCheckboxes(chkFileCreated, settings.getToggle("toggleFileCB"));
            ValidateToggleCheckboxes(chkEnumPurpose, settings.getToggle("toggleEnumPurpose"));
            ValidateToggleCheckboxes(chkEnumRH, settings.getToggle("toggleEnumRH"));
            ValidateToggleCheckboxes(chkEnumCreated, settings.getToggle("toggleEnumCB"));
            ValidateToggleCheckboxes(chkObjPurpose, settings.getToggle("toggleOBJPurpose"));
            ValidateToggleCheckboxes(chkObjRH, settings.getToggle("toggleOBJRH"));
            ValidateToggleCheckboxes(chkObjCreated, settings.getToggle("toggleOBJCB"));
            ValidateToggleCheckboxes(chkClassPurpose, settings.getToggle("toggleClassPurpose"));
            ValidateToggleCheckboxes(chkClassRH, settings.getToggle("toggleClassRH"));
            ValidateToggleCheckboxes(chkClassCreated, settings.getToggle("toggleClassCB"));
            ValidateToggleCheckboxes(chkMethodPurpose, settings.getToggle("toggleMethodPurpose"));
            ValidateToggleCheckboxes(chkMethodRH, settings.getToggle("toggleMethodRH"));
            ValidateToggleCheckboxes(chkMethodCreated, settings.getToggle("toggleMethodCB"));
            ValidateToggleCheckboxes(chkMethodEvents, settings.getToggle("toggleMethodEvents"));
            ValidateToggleCheckboxes(chkMethodExceptions, settings.getToggle("toggleMethodExcepts"));
            ValidateToggleCheckboxes(chkConstPurpose, settings.getToggle("toggleStructPurpose"));
            ValidateToggleCheckboxes(chkConstRH, settings.getToggle("toggleStructRH"));
            ValidateToggleCheckboxes(chkConstCreated, settings.getToggle("toggleStructCB"));
            ValidateToggleCheckboxes(chkConstEvents, settings.getToggle("toggleStructEvents"));
            ValidateToggleCheckboxes(chkConstExceptions, settings.getToggle("toggleStructExcepts"));
            setTextColor();
            SetXMLColors();
        }

        /// <METHOD name="ValidateToggleCheckboxes">
        /// <Purpose> 
        ///  Validates if the toggle checkbox is checked or not for the Toggle screen
        /// </Purpose>
        /// <Parameters> 
        ///     toggle(CheckBox):
        ///     enable(bool):
        /// </Parameters>
        /// </METHOD>
        public void ValidateToggleCheckboxes(CheckBox toggle, bool enable)
        {
            if (enable)
            {
                toggle.IsChecked = true;
            }
            else if (!enable)
            {

                toggle.IsChecked = false;
            }
        }

       /// <METHOD name="btnHomeScreen_Click">
       /// <Purpose> 
       ///  Generation horizontal menu button is clicked
       /// </Purpose>
       /// <Parameters> 
       ///     sender(object):
       ///     e(RoutedEventArgs):
       /// </Parameters>
       /// </METHOD>
        private void btnHomeScreen_Click(object sender, RoutedEventArgs e)
        {
            ChangeHomeGridVisibility(true);
            ChangeSettingsGridVisibility(false);

        }

        /// <METHOD name="btnSettingsScreen_Click">
        /// <Purpose> 
        ///  Settings horizontal menu button is clicked
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnSettingsScreen_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeHomeGridVisibility(false);
            LoadSettings();
        }

        /// <METHOD name="InitialGridVisibility">
        /// <Purpose> 
        ///  Sets the initial visibility of the GUI (Generation Screen)
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        /// 

        /// <METHOD name="ChangeSettingsGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeSettingsGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdSettingsGrid.Visibility = Visibility.Visible;
                grdSettingsGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdSettingsGrid.Visibility = Visibility.Hidden;
                grdSettingsGrid.IsEnabled = false;
            }
        }

        public void InitialGridVisibility()
        {
            ChangeHomeGridVisibility(true);
            ChangeSettingsGridVisibility(false);
            ChangeUpdateSelectedGridVisibility(false);
            ChangeSectionNameGridVisibility(false);
            ChangeSectionToggleGridVisibility(false);
            ChangeKeybindGridVisibility(false);
        }

        /// <METHOD name="ChangeOptionsStackEnable">
        /// <Purpose> 
        ///  Changes the enable status of the horizontal menu
        /// </Purpose>
        /// <Parameters> 
        ///     enable(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeOptionsStackEnable(bool enable)
        {
            if (enable)
            {
                stkScreenOptions.IsEnabled = true;
            }
            else if (!enable)
            {
                stkScreenOptions.IsEnabled = false;
            }
        }

        /// <METHOD name="ChangeHomeGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Home Grid
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeHomeGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdHomeGrid.Visibility = Visibility.Visible;
                grdHomeGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdHomeGrid.Visibility = Visibility.Hidden;
                grdHomeGrid.IsEnabled = false;
            }
        }

        /// <METHOD name="btnUpdateAll_Click">
        /// <Purpose> 
        ///  Full Template button is pressed on Generation screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnUpdateAll_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Comment Template Generated!");
            ChangeHomeGridVisibility(true);
            ChangeUpdateSelectedGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

        /// <METHOD name="btnUpdateSelected_Click">
        /// <Purpose> 
        ///  Partial Template button is pressed on Generation screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnUpdateSelected_Click(object sender, RoutedEventArgs e)
        {
            ChangeUpdateSelectedGridVisibility(true);
            ChangeHomeGridVisibility(false);
            ChangeOptionsStackEnable(false);
            LoadSettings();
        }

        /// <METHOD name="ChangeUpdateSelectedGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Partial Template screen
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeUpdateSelectedGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdUpdateSelectedGrid.Visibility = Visibility.Visible;
                grdUpdateSelectedGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdUpdateSelectedGrid.Visibility = Visibility.Hidden;
                grdUpdateSelectedGrid.IsEnabled = false;
            }
        }

        /// <METHOD name="btnRunUpdateSelected_Click">
        /// <Purpose> 
        ///  Run button is pressed on Partial Template screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private async void btnRunUpdateSelected_Click(object sender, RoutedEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null)
                return;

            Document actDocument = dte.ActiveDocument;
            TextDocument textDoc = (TextDocument)actDocument.Object("TextDocument");
            TextSelection selection = (TextSelection)actDocument.Selection;
            int originalLine = selection.ActivePoint.Line;
            int originalColumn = selection.ActivePoint.DisplayColumn;
            SyntaxTree syntaxTree = refreshTree(actDocument);

            if (chkNamespace.IsChecked == true)
            {
                allNamespace(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }
            if (chkClass.IsChecked == true)
            {
                allClasses(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }
            if (chkOBJProperties.IsChecked == true)
            {
                allObjAttributes(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }
            if (chkConstruct.IsChecked == true)
            {
                allStructs(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }
            if (chkEnum.IsChecked == true)
            {
                allEnums(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }
            if (chkMethod.IsChecked == true)
            {
                allMethods(dte, syntaxTree, textDoc);
                syntaxTree = refreshTree(actDocument);
                textDoc = refreshTextDoc(dte);
            }

            if (chkFile.IsChecked == true)
            {
                await InsertFlowerbox();
            }
            selection.MoveToLineAndOffset(originalLine, originalColumn);
            MessageBox.Show("Selected Template Sections Generated!");
            ChangeHomeGridVisibility(true);
            ChangeUpdateSelectedGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

        /// <METHOD name="btnCancelUpdateSelected_Click">
        /// <Purpose> 
        ///  Cancel button is pressed on Partial Template screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnCancelUpdateSelected_Click(object sender, RoutedEventArgs e)
        {
            ChangeHomeGridVisibility(true);
            ChangeUpdateSelectedGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

        /// <METHOD name="btnEditSectionNames_Click">
        /// <Purpose> 
        ///  Edit Section Names button is pressed on Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnEditSectionNames_Click(object sender, RoutedEventArgs e)
        {
            if (tagNames.initalized == false)
            {
                btnUpdateAllTags.IsEnabled = true;
                tagNames.initalized = true;
                tagNames.tagPurpose = txtPurpose.Text;
                tagNames.tagRevisionHistory = txtRevisionHistory.Text;
                tagNames.tagEventsHandled = txtEventHandled.Text;
                tagNames.tagEventRaised = txtEventRaised.Text;
                tagNames.tagExceptionCaught = txtExceptionCaught.Text;
                tagNames.tagExceptionThrown = txtExceptionThrown.Text;
                tagNames.tagCreatedBy = txtCreatedBy.Text;

            }

            ChangeSectionNameGridVisibility(true);
            ChangeSettingsGridVisibility(false);
            ChangeOptionsStackEnable(false);
            LoadSettings();
        }

        /// <METHOD name="ChangeSectionNameGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Edit Section Names screen
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeSectionNameGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdSectionNameGrid.Visibility = Visibility.Visible;
                grdSectionNameGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdSectionNameGrid.Visibility = Visibility.Hidden;
                grdSectionNameGrid.IsEnabled = false;
            }
        }

        /// <METHOD name="btnSaveSectionName_Click">
        /// <Purpose> 
        ///  Save Section Names button is pressed on the Edit Section Names screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnSaveSectionName_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeSectionNameGridVisibility(false);
            ChangeOptionsStackEnable(true);
            settings.setName("Purpose", txtPurpose.Text);
            settings.setName("Revise History", txtRevisionHistory.Text);
            settings.setName("Events Handled", txtEventHandled.Text);
            settings.setName("Event Raised", txtEventRaised.Text);
            settings.setName("Exception Caught", txtExceptionCaught.Text);
            settings.setName("Exception Thrown", txtExceptionThrown.Text);
            settings.setName("Created By", txtCreatedBy.Text);
            settings.setInitials(txtInitials.Text);
            settings.SaveSettings();
            btnUpdateAllTags_Click(btnUpdateAllTags, new RoutedEventArgs());
        }

        /// <METHOD name="btnCancelSectionName_Click">
        /// <Purpose> 
        ///  Cancel button is pressed on the Edit Section Names screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnCancelSectionName_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeSectionNameGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

       /// <METHOD name="btnToggleSections_Click">
       /// <Purpose> 
       ///  Customize Template button is pressed on the Settings Screen
       /// </Purpose>
       /// <Parameters> 
       ///     sender(object):
       ///     e(RoutedEventArgs):
       /// </Parameters>
       /// </METHOD>
        private void btnToggleSections_Click(object sender, RoutedEventArgs e)
        {
            ChangeSectionToggleGridVisibility(true);
            ChangeSettingsGridVisibility(false);
            ChangeOptionsStackEnable(false);
            LoadSettings();
        }

        /// <METHOD name="ChangeSectionToggleGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Customize Template screen
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeSectionToggleGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdSectionToggleGrid.Visibility = Visibility.Visible;
                grdSectionToggleGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdSectionToggleGrid.Visibility = Visibility.Hidden;
                grdSectionToggleGrid.IsEnabled = false;
            }
        }

        /// <METHOD name="btnSaveSectionToggle_Click">
        /// <Purpose> 
        ///  Save Template button is pressed on the Customize Template screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnSaveSectionToggle_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeSectionToggleGridVisibility(false);
            ChangeOptionsStackEnable(true);
            settings.setToggle("toggleNamespacePurpose", (bool)chkNamespacePurpose.IsChecked);
            settings.setToggle("toggleNamespaceRH", (bool)chkNamespaceRH.IsChecked);
            settings.setToggle("toggleNamespaceCB", (bool)chkNamespaceCreated.IsChecked);
            settings.setToggle("toggleFilePurpose", (bool)chkFilePurpose.IsChecked);
            settings.setToggle("toggleFileRH", (bool)chkFileRH.IsChecked);
            settings.setToggle("toggleFileCB", (bool)chkFileCreated.IsChecked);
            settings.setToggle("toggleEnumPurpose", (bool)chkEnumPurpose.IsChecked);
            settings.setToggle("toggleEnumRH", (bool)chkEnumRH.IsChecked);
            settings.setToggle("toggleEnumCB", (bool)chkEnumCreated.IsChecked);
            settings.setToggle("toggleOBJPurpose", (bool)chkObjPurpose.IsChecked);
            settings.setToggle("toggleOBJRH", (bool)chkObjRH.IsChecked);
            settings.setToggle("toggleOBJCB", (bool)chkObjCreated.IsChecked);
            settings.setToggle("toggleClassPurpose", (bool)chkClassPurpose.IsChecked);
            settings.setToggle("toggleClassRH", (bool)chkClassRH.IsChecked);
            settings.setToggle("toggleClassCB", (bool)chkClassCreated.IsChecked);
            settings.setToggle("toggleMethodPurpose", (bool)chkMethodPurpose.IsChecked);
            settings.setToggle("toggleMethodRH", (bool)chkMethodRH.IsChecked);
            settings.setToggle("toggleMethodCB", (bool)chkMethodCreated.IsChecked);
            settings.setToggle("toggleMethodEvents", (bool)chkMethodEvents.IsChecked);
            settings.setToggle("toggleMethodExcepts", (bool)chkMethodExceptions.IsChecked);
            settings.setToggle("toggleStructPurpose", (bool)chkConstPurpose.IsChecked);
            settings.setToggle("toggleStructRH", (bool)chkConstRH.IsChecked);
            settings.setToggle("toggleStructCB", (bool)chkConstCreated.IsChecked);
            settings.setToggle("toggleStructEvents", (bool)chkConstEvents.IsChecked);
            settings.setToggle("toggleStructExcepts", (bool)chkConstExceptions.IsChecked);
            settings.SaveSettings();
        }

       /// <METHOD name="btnCancelSectionToggle_Click">
       /// <Purpose> 
       ///  Cancel button is pressed on the Customize Template Screen
       /// </Purpose>
       /// <Parameters> 
       ///     sender(object):
       ///     e(RoutedEventArgs):
       /// </Parameters>
       /// </METHOD>
        private void btnCancelSectionToggle_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeSectionToggleGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

        /// <METHOD name="btnDefaultSettings_Click">
        /// <Purpose> 
        ///  Reset to Default Settings is pressed on the Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnDefaultSettings_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("You are about to reset all settings to default.\nAre you sure you want to continue?", "Reset Settings", MessageBoxButton.YesNo);
            if (MessageBoxResult.Yes == MessageBoxResult.Yes)
            {
                settings.DefaultSettings();
                SetXMLColors();
                btnUpdateAllTags_Click(btnUpdateAllTags, new RoutedEventArgs());
            }


        }

        /// <METHOD name="btnEditKeyBind_Click">
        /// <Purpose> 
        ///  Edit Color Settings is pressed on the Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnEditKeyBind_Click(object sender, RoutedEventArgs e)
        {
            ChangeKeybindGridVisibility(true);
            ChangeSettingsGridVisibility(false);
            ChangeOptionsStackEnable(false);
            LoadSettings();
        }

        /// <METHOD name="ChangeKeybindGridVisibility">
        /// <Purpose> 
        ///  Changes the visibility and enable status of the Edit Color Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     visibility(bool):
        /// </Parameters>
        /// </METHOD>
        public void ChangeKeybindGridVisibility(bool visibility)
        {
            if (visibility)
            {
                grdGeneralSettingsGrid.Visibility = Visibility.Visible;
                grdGeneralSettingsGrid.IsEnabled = true;
            }
            else if (!visibility)
            {
                grdGeneralSettingsGrid.Visibility = Visibility.Hidden;
                grdGeneralSettingsGrid.IsEnabled = false;
            }
        }

        /// <METHOD name="btnSaveKeybinds_Click">
        /// <Purpose> 
        ///  Save Colors button is pressed on the Edit Color Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private async void btnSaveKeybinds_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeKeybindGridVisibility(false);
            ChangeOptionsStackEnable(true);
            settings.setColor("OuterProperty", colorPickerProperty.Color.ToString());
            settings.setColor("InnerTag", colorPickerInner.Color.ToString());
            settings.setColor("OuterTag", colorPickerOuter.Color.ToString());
            settings.setInitials(txtInitials.Text);
            SetXMLColors();
            settings.SaveSettings();
            
        }

        /// <METHOD name="btnCancelKeybinds_Click">
        /// <Purpose> 
        ///  Cancel button is pressed on the Edit Color Settings screen
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnCancelKeybinds_Click(object sender, RoutedEventArgs e)
        {
            ChangeSettingsGridVisibility(true);
            ChangeKeybindGridVisibility(false);
            ChangeOptionsStackEnable(true);
        }

        /// <METHOD name="getDTE">
        /// <Purpose> 
        ///  Returns the DTE instance
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public DTE2 getDTE()
        {

            return COMETPackage.DTEInstance;

        }

        /// <METHOD name="getFontAndColorStorage">
        /// <Purpose> 
        ///  Returns the Font and Color Storage instance
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// </METHOD>
        public IVsFontAndColorStorage getFontAndColorStorage()
        {
            return COMETPackage.FontAndColorStorage;
        }

              /// <METHOD name="allComments">
       /// <Purpose> 
       ///  Generate comments for the Full Template
       /// </Purpose>
       /// <Parameters> 
       ///     sender(object):
       ///     e(RoutedEventArgs):
       /// </Parameters>
       /// </METHOD>
        private async void allComments(object sender, RoutedEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null)
                return;

            Document actDocument = dte.ActiveDocument;
            
            TextDocument textDoc = (TextDocument)actDocument.Object("TextDocument");
            TextSelection selection = (TextSelection)actDocument.Selection;
            int originalLine = selection.ActivePoint.Line;
            int originalColumn = selection.ActivePoint.DisplayColumn;

            // Initial parse
            SyntaxTree syntaxTree = refreshTree(actDocument);

            allNamespace(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            allClasses(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            allMethods(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            allEnums(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            allStructs(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            allObjAttributes(dte, syntaxTree, textDoc);
            syntaxTree = refreshTree(actDocument);
            textDoc = refreshTextDoc(dte);

            await InsertFlowerbox();
            
            // Restore selection to original spot
            selection.MoveToLineAndOffset(originalLine, originalColumn);
        }

        /// <METHOD name="refreshTextDoc">
        /// <Purpose> 
        ///  Refreshes the TextDocument object
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        /// </Parameters>
        /// </METHOD>
        private TextDocument refreshTextDoc(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Document actDocument = dte.ActiveDocument;
            TextDocument textDoc = (TextDocument)actDocument.Object("TextDocument");
            return textDoc;
        }

        /// <METHOD name="refreshTree">
        /// <Purpose> 
        ///  Refreshes the tree object
        /// </Purpose>
        /// <Parameters> 
        ///     activeDoc(Document):
        /// </Parameters>
        /// </METHOD>
        private SyntaxTree refreshTree(Document activeDoc)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            TextDocument textDoc = (TextDocument)activeDoc.Object("TextDocument");
            EditPoint start = textDoc.StartPoint.CreateEditPoint();
            string documentText = start.GetText(textDoc.EndPoint);

            CSharpParseOptions options = new CSharpParseOptions(default, DocumentationMode.Parse, SourceCodeKind.Regular, null);
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(documentText, options);
            return syntaxTree;
        }

        /// <METHOD name="allMethods">
        /// <Purpose> 
        ///  Locates all Methods and Constructors in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allMethods(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (dte?.ActiveDocument == null)
            {
                return;
            }

            var root = tree.GetRoot();
            var methodNodes = root.DescendantNodesAndSelf()
                                  .Where(node => node is MethodDeclarationSyntax || node is ConstructorDeclarationSyntax)
                                  .OrderByDescending(node => node.GetLocation().GetMappedLineSpan().StartLinePosition.Line)
                                  .ToList();

            foreach (var node in methodNodes)
            {
                if (node is MethodDeclarationSyntax methodNode)
                {
                    InsertMethodComment(methodNode, allowOverwrite: true);

                }
                else if (node is ConstructorDeclarationSyntax ctorNode)
                {
                    if (HasXmlDocumentation(ctorNode))
                        continue;

                    InsertConstructorComment(ctorNode);

                }
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="InsertConstructorComment">
        /// <Purpose> 
        ///  Inserts a comment for the constructor
        /// </Purpose>
        /// <Parameters> 
        ///     ctor(ConstructorDeclarationSyntax):
        /// </Parameters>
        /// </METHOD>
        public async void InsertConstructorComment(ConstructorDeclarationSyntax ctor)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null) return;
            var activeDoc = dte.ActiveDocument;
            string fileName = activeDoc.FullName;
            var textDoc = refreshTextDoc(dte);
            EditPoint editPoint = textDoc.CreateEditPoint();

            var span = ctor.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int columnOffset = span.StartLinePosition.Character;
            if (columnOffset == 0) columnOffset = 1;

            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, columnOffset)) return;

            string indent = new string(' ', columnOffset - 1);

            // ========== BUILD DYNAMIC CONTENT ==========

            string paramStr = string.Join(", ", ctor.ParameterList.Parameters.Select(p => p.Identifier.Text));
            string typeStr = string.Join("|", ctor.ParameterList.Parameters.Select(p => p.Type.ToString()));


            List<string> handledEvents = new List<string>();
          //  GetHandledEvents(ctor, fileName, out handledEvents);

            List<string> wiredEvents=new List<string>();
           // doesWireEventHandler(method, out wiredEvents);

            var thrownTypes = ctor.DescendantNodes().OfType<ThrowStatementSyntax>()
                .Select(t => t.Expression is ObjectCreationExpressionSyntax o ? o.Type.ToString() : "Exception").Distinct().ToList();

            var caughtTypes = ctor.DescendantNodes().OfType<CatchClauseSyntax>()
                .Select(c => c.Declaration?.Type.ToString() ?? "Exception").Distinct().ToList();

            // ========== EXTRACT OLD VALUES ==========


            var leadingTrivia = ctor.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            List<string> userContent = new List<string>();
            string newPreview = "";

            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();

                // === Step 1: Extract current (custom) section names ===
                string tagPurpose = sm.getName("Purpose");
                string tagRevise = sm.getName("Revise History");
                string tagCreatedBy = sm.getName("Created By");
                string tagWired = sm.getName("Event Raised");
                string tagHandled = sm.getName("Events Handled");
                string tagThrown = sm.getName("Exception Thrown");
                string tagCaught = sm.getName("Exception Caught");


                // === Step 2: Extract any existing user-entered content between tags ===
                string userPurpose = ExtractInnerTagContent(textDoc, row, "Purpose", tagPurpose);
                string userRevise = ExtractInnerTagContent(textDoc, row, "Revision History", tagRevise);
                string userCreatedBy = ExtractInnerTagContent(textDoc, row, "Created By", tagCreatedBy);
                string oldParameters = ExtractInnerTagContent(textDoc, row, "Parameters", "Parameters");
                string wired = ExtractInnerTagContent(textDoc, row, "Event Raised", tagWired);
                string handled = ExtractInnerTagContent(textDoc, row, "Events Handled", tagHandled);
                string thrown = ExtractInnerTagContent(textDoc, row, "Exception Thrown", tagThrown);
                string caught = ExtractInnerTagContent(textDoc, row, "Exception Caught", tagCaught);

                //add to list to pass to BuildMethodCommentText
                userContent.Add(userCreatedBy);
                userContent.Add(userPurpose);
                userContent.Add(oldParameters);
                userContent.Add(wired);
               userContent.Add(handled);
                userContent.Add(thrown);
                userContent.Add(caught);
                userContent.Add(userRevise);


                //  newPreview = BuildMethodCommentText(sm, method.Identifier.Text, indent, paramStr, typeStr,
                //wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent);

                string newParamStr = attachTypesToParams(paramStr, typeStr, indent);
                string correctFormatStr = removeNewLines(newParamStr);
                string newWiredStr = EventsUpdater.UpdateEvents(wired, wiredEvents, indent);
                string newHandledStr = EventsUpdater.UpdateEvents(handled, handledEvents, indent);
                string newThrown = EventsUpdater.UpdateEvents(thrown, thrownTypes, indent);
                string newCaught = EventsUpdater.UpdateEvents(caught, caughtTypes, indent);

                newPreview = BuildMethodCommentText(sm, ctor.Identifier.Text, indent, ParametersUpdater.UpdateParameters(oldParameters, correctFormatStr, indent), typeStr,
                        wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent, newWiredStr, newHandledStr, newThrown, newCaught);



                // Delete old comment
                int deleteLine = row;
                EditPoint commentStart = textDoc.CreateEditPoint();
                commentStart.MoveToLineAndOffset(deleteLine, 1);

                for (int i = deleteLine - 1; i >= 1; i--)
                {
                    string lineAbove = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                    if (!lineAbove.StartsWith("///"))
                        break;
                    commentStart.MoveToLineAndOffset(i, 1);
                }

                EditPoint commentEnd = textDoc.CreateEditPoint();
                commentEnd.MoveToLineAndOffset(row, columnOffset);
                commentStart.Delete(commentEnd);
                editPoint.Insert(newPreview);

            }
            else
            {
                newPreview = BuildMethodCommentText(sm, ctor.Identifier.Text, indent, paramStr, typeStr,
                    wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent, "", "", "", "");
                editPoint.Insert(newPreview);
                dte.Documents.SaveAll();
            }
        }

              /// <METHOD name="TryMoveToLineAndOffset">
       /// <Purpose> 
       ///  Sets the offset of the comment on the line
       /// </Purpose>
       /// <Parameters> 
       ///     textDoc(TextDocument):
       ///     editPoint(EditPoint):
       ///     row(int):
       ///     offset(int):
       /// </Parameters>
       /// <Exception_Caught> 
       ///     Exception:
       /// </Exception_Caught>
       /// </METHOD>
        private bool TryMoveToLineAndOffset(TextDocument textDoc, EditPoint editPoint, int row, int offset)
        {
            try
            {
                int totalLines = textDoc.EndPoint.Line;
                row = Math.Min(row, totalLines);
                if (row < 1) return false;

                string lineText = textDoc.CreateEditPoint().GetLines(row, row + 1);
                int maxOffset = lineText.Length;
                offset = Math.Min(offset, maxOffset);
                if (offset < 1) offset = 1;

                editPoint.MoveToLineAndOffset(row, offset);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR] Could not move to line {row}, offset {offset}: {ex.Message}");
                return false;
            }
        }

        /// <METHOD name="allClasses">
        /// <Purpose> 
        ///  Locates all Classes in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allClasses(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (dte?.ActiveDocument == null)
            {
                return;
            }

            var root = tree.GetRoot();
            var classNodes = root.DescendantNodesAndSelf().OfType<ClassDeclarationSyntax>()
                                  .OrderByDescending(c => c.GetLocation().GetMappedLineSpan().StartLinePosition.Line)
                                  .ToList();

            foreach (var classNode in classNodes)
            {
                InsertClassComment(classNode, allowOverwrite: true);
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="allEnums">
        /// <Purpose> 
        ///  Locates all Enums in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allEnums(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (dte?.ActiveDocument == null)
                return;

            var root = tree.GetRoot();
            var enumNodes = root.DescendantNodesAndSelf()
                .OfType<EnumDeclarationSyntax>()
                .OrderByDescending(e => e.GetLocation().GetMappedLineSpan().StartLinePosition.Line)
                .ToList();

            foreach (var enumNode in enumNodes)
            {
                InsertEnumComment(enumNode, allowOverwrite: true);
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="allStructs">
        /// <Purpose> 
        ///  Locates all Structs in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allStructs(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (dte?.ActiveDocument == null) return;

            var root = tree.GetRoot();
            var structNodes = root.DescendantNodesAndSelf()
                .OfType<StructDeclarationSyntax>()
                .OrderByDescending(s => s.GetLocation().GetMappedLineSpan().StartLinePosition.Line);

            foreach (var structNode in structNodes)
            {
                InsertStructComment(structNode);
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="allNamespace">
        /// <Purpose> 
        ///  Locate all Namespaces in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allNamespace(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (dte?.ActiveDocument == null)
                return;

            var root = tree.GetRoot();
            var namespaceNodes = root.DescendantNodesAndSelf().OfType<NamespaceDeclarationSyntax>()
                .OrderByDescending(n => n.GetLocation().GetMappedLineSpan().StartLinePosition.Line)
                .ToList();

            foreach (var ns in namespaceNodes)
            {
                InsertNamespaceComment(ns, allowOverwrite: true);
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="allObjAttributes">
        /// <Purpose> 
        ///  Locates all Object Attributes in the document and generates comments for them
        /// </Purpose>
        /// <Parameters> 
        ///     dte(DTE2):
        ///     tree(SyntaxTree):
        ///     textDoc(TextDocument):
        /// </Parameters>
        /// </METHOD>
        private async void allObjAttributes(DTE2 dte, SyntaxTree tree, TextDocument textDoc)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (dte?.ActiveDocument == null)
                return;

            var root = tree.GetRoot();
            var propertyNodes = root.DescendantNodes().OfType<PropertyDeclarationSyntax>()
                .OrderByDescending(p => p.GetLocation().GetMappedLineSpan().StartLinePosition.Line)
                .ToList();

            foreach (var property in propertyNodes)
            {
                InsertPropertyComment(property, allowOverwrite: true);
            }

            dte.Documents.SaveAll();
        }

        /// <METHOD name="HasXmlDocumentation">
        /// <Purpose> 
        ///  Checks to see if the node has XML documentation above it
        /// </Purpose>
        /// <Parameters> 
        ///     node(SyntaxNode):
        ///     expectedTag(string):
        /// </Parameters>
        /// </METHOD>
        private bool HasXmlDocumentation(SyntaxNode node, string expectedTag = null)
        {
            var trivia = node.GetLeadingTrivia();

            foreach (var t in trivia)
            {
                if (t.HasStructure && t.GetStructure() is DocumentationCommentTriviaSyntax doc)
                {
                    string commentText = doc.ToFullString();
                    if (expectedTag == null)
                        return true;

                    // Only consider it "existing" if it matches the expected XML tag
                    if (commentText.IndexOf($"<{expectedTag}", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }

            return false;
        }

        /// <METHOD name="InsertMethodComment">
        /// <Purpose> 
        ///  Inserts a comment for the method
        /// </Purpose>
        /// <Parameters> 
        ///     method(MethodDeclarationSyntax):
        ///     allowOverwrite(bool):
        /// </Parameters>
        /// </METHOD>
        public async void InsertMethodComment(MethodDeclarationSyntax method, bool allowOverwrite = false)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null) return;
            var activeDoc = dte.ActiveDocument;
            string fileName = activeDoc.FullName;
            var textDoc = refreshTextDoc(dte);
            EditPoint editPoint = textDoc.CreateEditPoint();

            var span = method.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int columnOffset = span.StartLinePosition.Character;
            if (columnOffset == 0) columnOffset = 1;

            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, columnOffset)) return;

            string indent = new string(' ', columnOffset - 1);

            // ========== BUILD DYNAMIC CONTENT ==========

            string paramStr = string.Join(", ", method.ParameterList.Parameters.Select(p => p.Identifier.Text));
            string typeStr = string.Join("|", method.ParameterList.Parameters.Select(p => p.Type.ToString()));


            List<string> handledEvents;
            GetHandledEvents(method, fileName, out handledEvents);

            List<string> wiredEvents;
            doesWireEventHandler(method, out wiredEvents);

            var thrownTypes = method.DescendantNodes().OfType<ThrowStatementSyntax>()
                .Select(t => t.Expression is ObjectCreationExpressionSyntax o ? o.Type.ToString() : "Exception").Distinct().ToList();

            var caughtTypes = method.DescendantNodes().OfType<CatchClauseSyntax>()
                .Select(c => c.Declaration?.Type.ToString() ?? "Exception").Distinct().ToList();

            // ========== EXTRACT OLD VALUES ==========
           

            var leadingTrivia = method.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            List<string> userContent = new List<string>();
            string newPreview="";

            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();

                // === Step 1: Extract current (custom) section names ===
                string tagPurpose = sm.getName("Purpose");
                string tagRevise = sm.getName("Revise History");
                string tagCreatedBy = sm.getName("Created By");
                string tagWired = sm.getName("Event Raised");
                string tagHandled = sm.getName("Events Handled");
                string tagThrown = sm.getName("Exception Thrown");
                string tagCaught = sm.getName("Exception Caught");


                // === Step 2: Extract any existing user-entered content between tags ===
                string userPurpose = ExtractInnerTagContent(textDoc, row, "Purpose", tagPurpose);
                string userRevise = ExtractInnerTagContent(textDoc, row, "Revision History", tagRevise);
                string userCreatedBy = ExtractInnerTagContent(textDoc, row, "Created By", tagCreatedBy);
                string oldParameters = ExtractInnerTagContent(textDoc, row, "Parameters", "Parameters");
                string wired = ExtractInnerTagContent(textDoc, row, "Event Raised", tagWired);
                string handled = ExtractInnerTagContent(textDoc, row, "Events Handled", tagHandled);
                string thrown = ExtractInnerTagContent(textDoc, row, "Exception Thrown", tagThrown);
                string caught = ExtractInnerTagContent(textDoc, row, "Exception Caught", tagCaught);

               //add to list to pass to BuildMethodCommentText
                userContent.Add(userCreatedBy);
                userContent.Add(userPurpose);
                userContent.Add(oldParameters);
                userContent.Add(wired);
                userContent.Add(handled);
                userContent.Add(thrown);
                userContent.Add(caught);
                userContent.Add(userRevise);

               
                  //  newPreview = BuildMethodCommentText(sm, method.Identifier.Text, indent, paramStr, typeStr,
                   //wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent);
              
                    string newParamStr = attachTypesToParams(paramStr, typeStr, indent);
                    string correctFormatStr = removeNewLines(newParamStr);
                    string newWiredStr = EventsUpdater.UpdateEvents(wired, wiredEvents, indent);
                    string newHandledStr = EventsUpdater.UpdateEvents(handled, handledEvents, indent);
                    string newThrown = EventsUpdater.UpdateEvents(thrown, thrownTypes, indent);
                    string newCaught = EventsUpdater.UpdateEvents(caught, caughtTypes, indent);

                newPreview = BuildMethodCommentText(sm, method.Identifier.Text, indent, ParametersUpdater.UpdateParameters(oldParameters,correctFormatStr,indent), typeStr,
                        wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent,newWiredStr,newHandledStr,newThrown,newCaught);
                    
                

                    // Delete old comment
                    int deleteLine = row;
                    EditPoint commentStart = textDoc.CreateEditPoint();
                    commentStart.MoveToLineAndOffset(deleteLine, 1);

                    for (int i = deleteLine - 1; i >= 1; i--)
                    {
                        string lineAbove = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                        if (!lineAbove.StartsWith("///"))
                            break;
                        commentStart.MoveToLineAndOffset(i, 1);
                    }

                    EditPoint commentEnd = textDoc.CreateEditPoint();
                    commentEnd.MoveToLineAndOffset(row, columnOffset);
                    commentStart.Delete(commentEnd);
                    editPoint.Insert(newPreview);

            }
            else
            {
                newPreview = BuildMethodCommentText(sm, method.Identifier.Text, indent, paramStr, typeStr,
                    wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent,"","","","");
                editPoint.Insert(newPreview);
                dte.Documents.SaveAll();
            }


                // string newComment = BuildMethodCommentText(sm, method.Identifier.Text, indent, paramStr, typeStr,
                // wiredEvents, handledEvents, thrownTypes, caughtTypes, fileName, userContent);

            
            
        }

        /// <METHOD name="BuildMethodCommentText">
        /// <Purpose> 
        ///  Builds the comment text for the method
        /// </Purpose>
        /// <Parameters> 
        ///     sm(SettingsManager):
        ///     methodName(string):
        ///     indent(string):
        ///     paramStr(string):
        ///     typeStr(string):
        ///     wired(List<string>):
        ///     handles(List<string>):
        ///     thrown(List<string>):
        ///     caught(List<string>):
        ///     documentPath(string):
        ///     userTexts(List<string>):
        ///     correctWired(string):
        ///     correctHandles(string):
        ///     correctThrown(string):
        ///     correctCaught(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        private string BuildMethodCommentText(SettingsManager sm, string methodName, string indent,
     string paramStr, string typeStr, List<string> wired, List<string> handles, List<string> thrown, List<string> caught, string documentPath,
     List<string> userTexts, string correctWired,string correctHandles, string correctThrown, string correctCaught)
        {
            char newLine = '\n';
            string comment = $@"{indent}/// <METHOD name=""{methodName}"">" + "\n";
            if (sm.getToggle("toggleMethodCB"))
                if (userTexts.Count == 0)
                {
                    comment += indent + $@"/// <{settings.getName("Created By")} name=""{settings.getInitials()}"" date=""{DateTime.Now:MM/dd/yyyy}""> {newLine}{indent}///{newLine}{indent}/// </{settings.getName("Created By")}>" + "\n";
                }
                else
                {
                    comment += indent + $@"/// <{settings.getName("Created By")} name=""{settings.getInitials()}"" date=""{DateTime.Now:MM/dd/yyyy}""> {convertXMLStringToMultiLine(userTexts[0],indent)} </{settings.getName("Created By")}>" + "\n";
                }
                    
            if (sm.getToggle("toggleMethodPurpose"))
                if (userTexts.Count == 0)
                {
                    comment += indent + $"/// <{settings.getName("Purpose")}> {newLine}{indent}///{newLine}{indent}/// </{settings.getName("Purpose")}>\n";
                }
                else
                {
                    comment += indent + $"/// <{settings.getName("Purpose")}> {convertXMLStringToMultiLine(userTexts[1],indent)} </{settings.getName("Purpose")}>\n";
                }

            if (!string.IsNullOrEmpty(paramStr))
            {
                if (paramStr.Contains("<Parameters>"))
                { 
                    if(paramStr.Equals("/// <Parameters>\n\n       /// </Parameters>\n"))
                    {
                        comment += indent + $"/// <Parameters> {newLine}{indent}///{newLine}{indent}/// </Parameters>\n";
                    }
                    else
                    {
                        comment += indent + paramStr;
                    }
                       
                }
                else
                {


                    if (paramStr.IsNullOrWhiteSpace())
                    {
                        comment += indent + $"/// <Parameters> {newLine}{indent}///{newLine}{indent}/// </Parameters>\n";
                    }
                    else
                    {
                        paramStr = attachTypesToParams(paramStr, typeStr, indent);
                        comment += indent + $"/// <Parameters> {paramStr} </Parameters>\n";
                    }

                }
            }
            else
            {
                
                comment += indent + $"/// <Parameters> {newLine}{indent}///{newLine}{indent}/// </Parameters>\n";
            }

            if (sm.getToggle("toggleMethodEvents") && (wired.Any() || handles.Any()))
            {
                if (userTexts.Count == 0)
                {
                    if (wired.Any())
                        comment += indent + $"/// <{sm.getName("Event Raised")}> {convertEventsToXML(wired, indent)} </{sm.getName("Event Raised")}>\n";
                    if (handles.Any())
                        comment += indent + $"/// <{sm.getName("Events Handled")}> {convertEventsToXML(handles, indent)} </{sm.getName("Events Handled")}>  \n";

                }
                else
                {
                    if (wired.Any())
                        comment += indent + $"/// <{sm.getName("Event Raised")}> \n{correctWired}{indent}/// </{sm.getName("Event Raised")}>\n";
                    if (handles.Any())
                        comment += indent + $"/// <{sm.getName("Events Handled")}> \n{correctHandles}{indent}/// </{sm.getName("Events Handled")}>\n";
                }

            }
            if (sm.getToggle("toggleMethodExcepts") && (thrown.Any() || caught.Any()))
            {
                if (userTexts.Count == 0)
                {
                    if (thrown.Any())
                        comment += indent + $"/// <{sm.getName("Exception Thrown")}> {convertExecptionToXML(thrown, indent)} </{sm.getName("Exception Thrown")}>\n";
                    if (caught.Any())
                        comment += indent + $"/// <{sm.getName("Exception Caught")}> {convertExecptionToXML(caught, indent)} </{sm.getName("Exception Caught")}>\n";
                }
                else
                {
                    if (thrown.Any())
                        comment += indent + $"/// <{sm.getName("Exception Thrown")}> \n {correctThrown}{indent}/// </{sm.getName("Exception Thrown")}>\n";
                    if (caught.Any())
                        comment += indent + $"/// <{sm.getName("Exception Caught")}> \n {correctCaught}{indent}/// </{sm.getName("Exception Caught")}>\n";
                }
                    
            }
            if (sm.getToggle("toggleMethodRH"))
            {
                if (userTexts.Count == 0)
                {
                    comment += indent + $"/// <{settings.getName("Revise History")}> {newLine}{indent}///{newLine}{indent}/// </{settings.getName("Revise History")}>\n";
                }
                else
                {
                    comment += indent + $"/// <{settings.getName("Revise History")}>{convertXMLStringToMultiLine(userTexts[7],indent)}</{settings.getName("Revise History")}>\n"; ;
                }
            }
               

            comment += indent + "/// </METHOD>\n" + indent;
            return comment;
        }

        /// <METHOD name="NormalizeComment">
        /// <Purpose> 
        ///  Replaces all spaces, tabs, new lines and carriage returns with empty strings
        /// </Purpose>
        /// <Parameters> 
        ///     comment(string):
        /// </Parameters>
        /// </METHOD>
        private string NormalizeComment(string comment)
        {
            return comment.Replace(" ", "").Replace("\t", "").Replace("\n", "").Replace("\r", "").ToLowerInvariant();
        }

        /// <METHOD name="InsertPropertyComment">
        /// <Purpose> 
        ///  Inserts the Object Property comment
        /// </Purpose>
        /// <Parameters> 
        ///     prop(PropertyDeclarationSyntax):
        ///     allowOverwrite(bool):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public void InsertPropertyComment(PropertyDeclarationSyntax prop, bool allowOverwrite = false)
        {
            SettingsManager sm = new SettingsManager();
            var dte = getDTE();
            var textDoc = refreshTextDoc(dte);

            var span = prop.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int offset = span.StartLinePosition.Character;
            if (offset == 0) offset = 1;

            EditPoint editPoint = textDoc.CreateEditPoint();
            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, offset))
                return;

            string indent = new string(' ', offset - 1);

            string labelCreatedBy = sm.getName("Created By");
            string labelPurpose = sm.getName("Purpose");
            string labelRevise = sm.getName("Revise History");

            string userCreatedBy = ExtractInnerTagContentHelper(textDoc, row, "Created By", labelCreatedBy);
            string userPurpose = ExtractInnerTagContentHelper(textDoc, row, "Purpose", labelPurpose);
            string userRevise = ExtractInnerTagContentHelper(textDoc, row, "Revision History", labelRevise);

            string comment = indent + $"/// <PROPERTY name=\"{prop.Identifier.Text}\">\n";

            if (sm.getToggle("toggleOBJCB"))
            {
                comment += indent + $"/// <{labelCreatedBy}>\n";
                comment += convertDynamicToXMLFormatHelper(userCreatedBy, indent) + "\n";
                comment += indent + $"/// </{labelCreatedBy}>\n";
            }

            if (sm.getToggle("toggleOBJPurpose"))
            {
                comment += indent + $"/// <{labelPurpose}>\n";
                comment += convertDynamicToXMLFormatHelper(userPurpose, indent) + "\n";
                comment += indent + $"/// </{labelPurpose}>\n";
            }

            if (sm.getToggle("toggleOBJRH"))
            {
                comment += indent + $"/// <{labelRevise}>\n";
                comment += convertDynamicToXMLFormatHelper(userRevise, indent) + "\n";
                comment += indent + $"/// </{labelRevise}>\n";
            }

            comment += indent + "/// </PROPERTY>\n" + indent;

            var leadingTrivia = prop.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            bool needsUpdate = true;

            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();
                if (NormalizeComment(existing) == NormalizeComment(comment))
                {
                    needsUpdate = false;
                }
                else
                {
                    int deleteOffsetStart = span.StartLinePosition.Line;
                    var commentLineStart = textDoc.CreateEditPoint();
                    commentLineStart.MoveToLineAndOffset(deleteOffsetStart, 1);

                    for (int i = deleteOffsetStart - 1; i >= 1; i--)
                    {
                        string lineText = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                        if (!lineText.StartsWith("///"))
                            break;
                        commentLineStart.MoveToLineAndOffset(i, 1);
                    }

                    EditPoint commentEnd = commentLineStart.CreateEditPoint();
                    commentEnd.MoveToLineAndOffset(row, offset);
                    commentLineStart.Delete(commentEnd);
                }
            }

            if (needsUpdate)
            {
                editPoint.Insert(comment);
                dte.Documents.SaveAll();
            }
        }

        /// <METHOD name="ExtractInnerTagContent">
        /// <Purpose> 
        ///  Pulls the contents of the inner tag
        /// </Purpose>
        /// <Parameters> 
        ///     textDoc(TextDocument):
        ///     startLine(int):
        ///     fallbackTag(string):
        ///     currentTag(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     content:
        ///     content:
        /// </Event_Raised>
        /// </METHOD>
        private string ExtractInnerTagContent(TextDocument textDoc, int startLine, string fallbackTag, string currentTag)
        {
            string content = "";
            bool saving = false;
            List<string> stringList = new List<string>();
            for (int i = startLine - 1; i >= 1; i--)
            {
                string line = textDoc.CreateEditPoint().GetLines(i, i + 1);

                string lower = line;

                if (!lower.Contains("///"))
                {
                    return "";
                    break;

                }


                if (saving)
                {

                    if (lower.Contains("<" + currentTag) && lower.Contains(">"))
                    {
                        stringList.Add(lower);

                        for (int j = stringList.Count - 2; j >= 1;j--)
                        {
                            content += stringList[j];
                        }

                        return content;

                    }
                    else
                    {
                        stringList.Add(lower);
                    }
                }


                if (lower.Contains("</" + currentTag + ">"))
                {
                    if(lower.Contains("</" + currentTag + ">") && lower.Contains("<" + currentTag))
                    {
                        return lower;
                    }
                    saving = true;
                    stringList.Add(lower);

                }
                if (!lower.Contains("</" + currentTag + ">")) continue;

            }

            for (int i = stringList.Count - 1; i >= 0;)
            {
                content += stringList[i];
            }

            return content;
        }

        /// <METHOD name="ExtractInnerTagContentHelper">
        /// <Purpose> 
        ///  Pulls the contents of the inner tag
        /// </Purpose>
        /// <Parameters> 
        ///     textDoc(TextDocument):
        ///     startLine(int):
        ///     fallbackTag(string):
        ///     currentTag(string):
        /// </Parameters>
        /// </METHOD>
        private string ExtractInnerTagContentHelper(TextDocument textDoc, int startLine, string fallbackTag, string currentTag)
        {
            var lines = new List<string>();
            bool saving = false;

            for (int i = startLine - 1; i >= 1; i--)
            {
                string line = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();

                if (!line.StartsWith("///"))
                    break;

                if (line.Contains($"</{currentTag}>") || line.Contains($"</{fallbackTag}>"))
                {
                    saving = true;
                    continue;
                }

                if (line.Contains($"<{currentTag}") || line.Contains($"<{fallbackTag}"))
                {
                    break;
                }

                if (saving)
                {
                    string content = line.Replace("///", "").Trim();
                    lines.Insert(0, content);
                }
            }

            return string.Join("\n", lines);
        }

        /// <METHOD name="convertDynamicToXMLFormatHelper">
        /// <Purpose> 
        ///  Formats the user text to XML format
        /// </Purpose>
        /// <Parameters> 
        ///     userText(string):
        ///     indent(string):
        /// </Parameters>
        /// </METHOD>
        private string convertDynamicToXMLFormatHelper(string userText, string indent)
        {
            if (string.IsNullOrWhiteSpace(userText))
                return indent + "///";

            var lines = userText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            return string.Join("\n", lines.Select(line =>
            {
                string trimmed = line.Trim();
                return trimmed.StartsWith("///") ? indent + trimmed : indent + "/// " + trimmed;
            }));
        }

        /// <METHOD name="InsertClassComment">
        /// <Purpose> 
        ///  Inserts the Class comment
        /// </Purpose>
        /// <Parameters> 
        ///     cls(ClassDeclarationSyntax):
        ///     allowOverwrite(bool):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public async void InsertClassComment(ClassDeclarationSyntax cls, bool allowOverwrite = false)
        {
            SettingsManager sm = new SettingsManager();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var dte = getDTE();
            var textDoc = refreshTextDoc(dte);

            var span = cls.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int offset = span.StartLinePosition.Character;
            if (offset == 0) offset = 1;

            EditPoint editPoint = textDoc.CreateEditPoint();
            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, offset))
                return;

            string indent = new string(' ', offset - 1);

            // === Step 1: Extract current (custom) section names ===
            string labelCreatedBy = sm.getName("Created By");
            string labelPurpose = sm.getName("Purpose");
            string labelRevise = sm.getName("Revise History");

            // === Step 2: Extract any existing user-entered content between tags ===
            string userCreatedBy = ExtractInnerTagContentHelper(textDoc, row, "Created By", labelCreatedBy);
            string userPurpose = ExtractInnerTagContentHelper(textDoc, row, "Purpose", labelPurpose);
            string userRevise = ExtractInnerTagContentHelper(textDoc, row, "Revision History", labelRevise);

            // === Step 3: Build comment with user text ===
            string comment = indent + $"/// <CLASS name=\"{cls.Identifier.Text}\">\n";

            if (sm.getToggle("toggleClassCB"))
            {
                comment += indent + $"/// <{labelCreatedBy}>\n";
                comment += convertDynamicToXMLFormatHelper(userCreatedBy, indent) + "\n";
                comment += indent + $"/// </{labelCreatedBy}>\n";
            }

            if (sm.getToggle("toggleClassPurpose"))
            {
                comment += indent + $"/// <{labelPurpose}>\n";
                comment += convertDynamicToXMLFormatHelper(userPurpose, indent) + "\n";
                comment += indent + $"/// </{labelPurpose}>\n";
            }

            if (sm.getToggle("toggleClassRH"))
            {
                comment += indent + $"/// <{labelRevise}>\n";
                comment += convertDynamicToXMLFormatHelper(userRevise, indent) + "\n";
                comment += indent + $"/// </{labelRevise}>\n";
            }

            comment += indent + "/// </CLASS>\n" + indent;

            // === Step 4: Remove old comment if it exists and differs ===
            var leadingTrivia = cls.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            bool needsUpdate = true;

            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();
                if (NormalizeComment(existing) == NormalizeComment(comment))
                {
                    needsUpdate = false;
                }
                else
                {
                    // Delete old comment
                    int deleteLine = row;
                    EditPoint commentStart = textDoc.CreateEditPoint();
                    commentStart.MoveToLineAndOffset(deleteLine, 1);

                    for (int i = deleteLine - 1; i >= 1; i--)
                    {
                        string lineAbove = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                        if (!lineAbove.StartsWith("///")) break;
                        commentStart.MoveToLineAndOffset(i, 1);
                    }

                    EditPoint commentEnd = textDoc.CreateEditPoint();
                    commentEnd.MoveToLineAndOffset(row, offset);
                    commentStart.Delete(commentEnd);
                }
            }

            if (needsUpdate)
            {
                editPoint.Insert(comment);
                dte.Documents.SaveAll();
            }
        }

        /// <METHOD name="InsertEnumComment">
        /// <Purpose> 
        ///  Inserts Enum comment
        /// </Purpose>
        /// <Parameters> 
        ///     en(EnumDeclarationSyntax):
        ///     allowOverwrite(bool):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public void InsertEnumComment(EnumDeclarationSyntax en, bool allowOverwrite = false)
        {
            SettingsManager sm = new SettingsManager();
            var dte = getDTE();
            if (dte?.ActiveDocument == null) return;

            var textDoc = refreshTextDoc(dte);
            EditPoint editPoint = textDoc.CreateEditPoint();

            var span = en.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int offset = span.StartLinePosition.Character;
            if (offset == 0) offset = 1;

            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, offset)) return;
            string indent = new string(' ', offset - 1);

            string labelCB = sm.getName("Created By");
            string labelPurpose = sm.getName("Purpose");
            string labelRH = sm.getName("Revise History");

            string userCB = ExtractInnerTagContentHelper(textDoc, row, "Created By", labelCB);
            string userPurpose = ExtractInnerTagContentHelper(textDoc, row, "Purpose", labelPurpose);
            string userRH = ExtractInnerTagContentHelper(textDoc, row, "Revision History", labelRH);

            string comment = indent + $"/// <ENUM name=\"{en.Identifier.Text}\">\n";
            if (sm.getToggle("toggleEnumCB"))
            {
                comment += indent + $"/// <{labelCB}>\n";
                comment += convertDynamicToXMLFormatHelper(userCB, indent) + "\n";
                comment += indent + $"/// </{labelCB}>\n";
            }
            if (sm.getToggle("toggleEnumPurpose"))
            {
                comment += indent + $"/// <{labelPurpose}>\n";
                comment += convertDynamicToXMLFormatHelper(userPurpose, indent) + "\n";
                comment += indent + $"/// </{labelPurpose}>\n";
            }
            if (sm.getToggle("toggleEnumRH"))
            {
                comment += indent + $"/// <{labelRH}>\n";
                comment += convertDynamicToXMLFormatHelper(userRH, indent) + "\n";
                comment += indent + $"/// </{labelRH}>\n";
            }
            comment += indent + "/// </ENUM>\n" + indent;

            var leadingTrivia = en.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            bool needsUpdate = true;

            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();
                if (NormalizeComment(existing) == NormalizeComment(comment))
                    needsUpdate = false;
                else
                {
                    int deleteOffsetStart = span.StartLinePosition.Line;
                    var commentLineStart = textDoc.CreateEditPoint();
                    commentLineStart.MoveToLineAndOffset(deleteOffsetStart, 1);
                    for (int i = deleteOffsetStart - 1; i >= 1; i--)
                    {
                        string text = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                        if (!text.StartsWith("///")) break;
                        commentLineStart.MoveToLineAndOffset(i, 1);
                    }
                    EditPoint commentEnd = commentLineStart.CreateEditPoint();
                    commentEnd.MoveToLineAndOffset(row, offset);
                    commentLineStart.Delete(commentEnd);
                }
            }

            if (needsUpdate)
            {
                editPoint.Insert(comment);
                dte.Documents.SaveAll();
            }
        }

        /// <METHOD name="InsertStructComment">
        /// <Purpose> 
        ///  Inserts Struct comment
        /// </Purpose>
        /// <Parameters> 
        ///     structNode(StructDeclarationSyntax):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public async void InsertStructComment(StructDeclarationSyntax structNode)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null) return;

            TextDocument textDoc = refreshTextDoc(dte);
            EditPoint editPoint = textDoc.CreateEditPoint();

            var span = structNode.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int offset = span.StartLinePosition.Character;
            if (offset == 0) offset = 1;

            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, offset)) return;

            string indent = new string(' ', offset - 1);
            var sm = new SettingsManager();

            string labelCreatedBy = sm.getName("Created By");
            string labelPurpose = sm.getName("Purpose");
            string labelRevise = sm.getName("Revise History");

            string userCreatedBy = ExtractInnerTagContentHelper(textDoc, row, "Created By", labelCreatedBy);
            string userPurpose = ExtractInnerTagContentHelper(textDoc, row, "Purpose", labelPurpose);
            string userRevise = ExtractInnerTagContentHelper(textDoc, row, "Revision History", labelRevise);

            string comment = indent + $"/// <STRUCTURE name=\"{structNode.Identifier.Text}\">\n";

            if (sm.getToggle("toggleStructCB"))
            {
                comment += indent + $"/// <{labelCreatedBy}>\n";
                comment += convertDynamicToXMLFormatHelper(userCreatedBy, indent) + "\n";
                comment += indent + $"/// </{labelCreatedBy}>\n";
            }

            if (sm.getToggle("toggleStructPurpose"))
            {
                comment += indent + $"/// <{labelPurpose}>\n";
                comment += convertDynamicToXMLFormatHelper(userPurpose, indent) + "\n";
                comment += indent + $"/// </{labelPurpose}>\n";
            }

            if (sm.getToggle("toggleStructRH"))
            {
                comment += indent + $"/// <{labelRevise}>\n";
                comment += convertDynamicToXMLFormatHelper(userRevise, indent) + "\n";
                comment += indent + $"/// </{labelRevise}>\n";
            }

            comment += indent + "/// </STRUCTURE>\n" + indent;

            var leadingTrivia = structNode.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);

            bool needsUpdate = true;
            if (docTrivia != default)
            {
                string existing = docTrivia.ToFullString().Trim();
                if (NormalizeComment(existing) == NormalizeComment(comment))
                {
                    needsUpdate = false;
                }
                else
                {
                    int deleteOffsetStart = span.StartLinePosition.Line;
                    var commentLineStart = textDoc.CreateEditPoint();
                    commentLineStart.MoveToLineAndOffset(deleteOffsetStart, 1);

                    for (int i = deleteOffsetStart - 1; i >= 1; i--)
                    {
                        string lineText = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                        if (!lineText.StartsWith("///")) break;
                        commentLineStart.MoveToLineAndOffset(i, 1);
                    }

                    EditPoint commentEnd = commentLineStart.CreateEditPoint();
                    commentEnd.MoveToLineAndOffset(row, offset);
                    commentLineStart.Delete(commentEnd);
                }
            }

            if (needsUpdate)
            {
                editPoint.Insert(comment);
                dte.Documents.SaveAll();
            }
        }

        /// <METHOD name="InsertNamespaceComment">
        /// <Purpose> 
        ///  Inserts Namespace comment
        /// </Purpose>
        /// <Parameters> 
        ///     ns(NamespaceDeclarationSyntax):
        ///     allowOverwrite(bool):
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public void InsertNamespaceComment(NamespaceDeclarationSyntax ns, bool allowOverwrite = false)
        {
            var dte = getDTE();
            var textDoc = refreshTextDoc(dte);

            var span = ns.GetLocation().GetMappedLineSpan();
            int row = span.StartLinePosition.Line + 1;
            int offset = span.StartLinePosition.Character;
            if (offset == 0) offset = 1;

            EditPoint editPoint = textDoc.CreateEditPoint();
            if (!TryMoveToLineAndOffset(textDoc, editPoint, row, offset))
                return;

            string indent = new string(' ', offset - 1);

            string labelCreatedBy = settings.getName("Created By");
            string labelPurpose = settings.getName("Purpose");
            string labelRevise = settings.getName("Revise History");

            string userCreatedBy = ExtractInnerTagContentHelper(textDoc, row, "Created By", labelCreatedBy);
            string userPurpose = ExtractInnerTagContentHelper(textDoc, row, "Purpose", labelPurpose);
            string userRevise = ExtractInnerTagContentHelper(textDoc, row, "Revise History", labelRevise);

            string comment = indent + $"/// <NAMESPACE name=\"{ns.Name}\">\n";

            if (settings.getToggle("toggleNamespaceCB"))
            {
                comment += indent + $"/// <{labelCreatedBy}>\n";
                comment += convertDynamicToXMLFormatHelper(userCreatedBy, indent) + "\n";
                comment += indent + $"/// </{labelCreatedBy}>\n";
            }

            if (settings.getToggle("toggleNamespacePurpose"))
            {
                comment += indent + $"/// <{labelPurpose}>\n";
                comment += convertDynamicToXMLFormatHelper(userPurpose, indent) + "\n";
                comment += indent + $"/// </{labelPurpose}>\n";
            }

            if (settings.getToggle("toggleNamespaceRH"))
            {
                comment += indent + $"/// <{labelRevise}>\n";
                comment += convertDynamicToXMLFormatHelper(userRevise, indent) + "\n";
                comment += indent + $"/// </{labelRevise}>\n";
            }

            comment += indent + "/// </NAMESPACE>\n" + indent;

            var leadingTrivia = ns.GetLeadingTrivia();
            var docTrivia = leadingTrivia.FirstOrDefault(t => t.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia);
            if (docTrivia != default)
            {
                int deleteOffsetStart = span.StartLinePosition.Line;
                var commentLineStart = textDoc.CreateEditPoint();
                commentLineStart.MoveToLineAndOffset(deleteOffsetStart, 1);

                for (int i = deleteOffsetStart - 1; i >= 1; i--)
                {
                    string currentLine = textDoc.CreateEditPoint().GetLines(i, i + 1).Trim();
                    if (!currentLine.StartsWith("///") || currentLine.Contains("<FLOWERBOX"))
                        break;
                    commentLineStart.MoveToLineAndOffset(i, 1);
                }

                EditPoint commentEnd = commentLineStart.CreateEditPoint();
                commentEnd.MoveToLineAndOffset(row, offset);
                commentLineStart.Delete(commentEnd);
            }

            editPoint.Insert(comment);
            dte.Documents.SaveAll();
        }

        /// <METHOD name="InsertFlowerbox">
        /// <Purpose> 
        ///  Insert a flowerbox comment at the top of the file
        /// </Purpose>
        /// <Parameters> 
        ///
        /// </Parameters>
        /// <Event_Raised> 
        ///     comment:
        /// </Event_Raised>
        /// </METHOD>
        public async Task InsertFlowerbox()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var dte = getDTE();
            if (dte?.ActiveDocument == null) return;

            Document actDocument = dte.ActiveDocument;
            TextDocument textDoc = (TextDocument)actDocument.Object("TextDocument");
            EditPoint editPoint = textDoc.CreateEditPoint();
            string documentText = editPoint.GetText(textDoc.EndPoint);

            var lines = documentText.Split('\n');
            int flowerboxStart = -1;
            int flowerboxEnd = -1;

            // Locate existing flowerbox bounds
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("<FLOWERBOX"))
                {
                    flowerboxStart = i;
                    for (int j = i; j < lines.Length; j++)
                    {
                        if (lines[j].Contains("</FLOWERBOX>"))
                        {
                            flowerboxEnd = j;
                            break;
                        }
                    }
                    break;
                }
            }

            string existingCB = "";
            string existingPurpose = "";
            string existingRevise = "";

            // Extract old content if it exists
            if (flowerboxStart != -1 && flowerboxEnd != -1)
            {
                for (int i = flowerboxStart; i <= flowerboxEnd; i++)
                {
                    string line = lines[i].Trim();

                    if (line.Contains($"<{settings.getName("Created By")}>"))
                    {
                        var collectedLines = new List<string>();
                        for (int j = i + 1; j <= flowerboxEnd; j++)
                        {
                            if (lines[j].Contains($"</{settings.getName("Created By")}>"))
                                break;
                            collectedLines.Add(lines[j].TrimEnd());
                        }
                        existingCB = string.Join("\n", collectedLines);
                    }

                    if (line.Contains($"<{settings.getName("Purpose")}>"))
                    {
                        var collectedLines = new List<string>();
                        for (int j = i + 1; j <= flowerboxEnd; j++)
                        {
                            if (lines[j].Contains($"</{settings.getName("Purpose")}>"))
                                break;
                            collectedLines.Add(lines[j].TrimEnd());
                        }
                        existingPurpose = string.Join("\n", collectedLines);
                    }

                    if (line.Contains($"<{settings.getName("Revise History")}>"))
                    {
                        var collectedLines = new List<string>();
                        for (int j = i + 1; j <= flowerboxEnd; j++)
                        {
                            if (lines[j].Contains($"</{settings.getName("Revise History")}>"))
                                break;
                            collectedLines.Add(lines[j].TrimEnd());
                        }
                        existingRevise = string.Join("\n", collectedLines);
                    }
                }

                // Remove old flowerbox
                EditPoint startPoint = textDoc.CreateEditPoint();
                EditPoint endPoint = textDoc.CreateEditPoint();
                startPoint.MoveToLineAndOffset(flowerboxStart + 1, 1);
                endPoint.MoveToLineAndOffset(flowerboxEnd + 2, 1);
                startPoint.Delete(endPoint);
            }

            // Determine where to insert
            SyntaxTree tree = CSharpSyntaxTree.ParseText(textDoc.CreateEditPoint().GetText(textDoc.EndPoint));
            var root = tree.GetRoot();
            var lastUsing = root.DescendantNodes().OfType<UsingDirectiveSyntax>().LastOrDefault();
            int insertLine = 1;
            if (lastUsing != null)
                insertLine = lastUsing.GetLocation().GetMappedLineSpan().EndLinePosition.Line + 2;

            string fileName = Path.GetFileName(actDocument.FullName);
            string indent = "";
            string comment = $@"{indent}/// <FLOWERBOX file=""{fileName}"">" + "\n";

            if (settings.getToggle("toggleFileCB"))
            {
                comment += indent + $"/// <{settings.getName("Created By")}>\n";
                comment += convertDynamicToXMLFormatHelper(existingCB, indent) + "\n";
                comment += indent + $"/// </{settings.getName("Created By")}>\n";
            }

            if (settings.getToggle("toggleFilePurpose"))
            {
                comment += indent + $"/// <{settings.getName("Purpose")}>\n";
                comment += convertDynamicToXMLFormatHelper(existingPurpose, indent) + "\n";
                comment += indent + $"/// </{settings.getName("Purpose")}>\n";
            }

            if (settings.getToggle("toggleFileRH"))
            {
                comment += indent + $"/// <{settings.getName("Revise History")}>\n";
                comment += convertDynamicToXMLFormatHelper(existingRevise, indent) + "\n";
                comment += indent + $"/// </{settings.getName("Revise History")}>\n";
            }

            comment += indent + "/// </FLOWERBOX>\n" + indent;

            // Insert new flowerbox
            EditPoint insertPoint = textDoc.CreateEditPoint();
            insertPoint.MoveToLineAndOffset(insertLine, 1);
            insertPoint.Insert(comment);

            dte.Documents.SaveAll();
        }

        /// <METHOD name="GetHandledEvents">
        /// <Purpose> 
        ///  Confirms if the event handler method is wired to any events
        /// </Purpose>
        /// <Parameters> 
        ///     eventHandlerMethod(MethodDeclarationSyntax):
        ///     sourceFilePath(string):
        ///     handledEvents(List<string>):
        /// </Parameters>
        /// </METHOD>
        public bool GetHandledEvents(MethodDeclarationSyntax eventHandlerMethod, string sourceFilePath, out List<string> handledEvents)
        {
            // Initialize the output list
            handledEvents = new List<string>();

            // Get the event handler method's name (e.g., "Button1_Click")
            string methodName = eventHandlerMethod.Identifier.Text;

            // Get the syntax tree of the source file from the MethodDeclarationSyntax
            SyntaxTree sourceTree = eventHandlerMethod.SyntaxTree;

            // Determine the designer file path (e.g., replace "Form1.cs" with "Form1.Designer.cs")
            string designerFilePath = Path.ChangeExtension(sourceFilePath, ".Designer.cs");

            // List of syntax trees to analyze (starting with the source file)
            List<SyntaxTree> treesToAnalyze = new List<SyntaxTree> { sourceTree };

            // Check if the designer file exists, and if so, parse it
            if (File.Exists(designerFilePath))
            {
                string designerCode = File.ReadAllText(designerFilePath);
                SyntaxTree designerTree = CSharpSyntaxTree.ParseText(designerCode);
                treesToAnalyze.Add(designerTree);
            }

            // Analyze each syntax tree for event subscriptions
            foreach (var tree in treesToAnalyze)
            {
                var root = tree.GetCompilationUnitRoot();

                // Find all += assignment expressions
                var assignments = root.DescendantNodes()
                    .OfType<AssignmentExpressionSyntax>()
                    .Where(a => a.OperatorToken.Text == "+=");

                foreach (var assignment in assignments)
                {
                    // Check if the right-hand side references the event handler
                    if (IsEventHandlerReference(assignment.Right, methodName))
                    {
                        // Add the event (left-hand side) to the list
                        handledEvents.Add(assignment.Left.ToString());
                    }
                }
            }

            // Return true if any events were found, false otherwise
            return handledEvents.Any();
        }

        /// <METHOD name="IsEventHandlerReference">
        /// <Purpose> 
        ///  Checks if the expression is a reference to the event handler method
        /// </Purpose>
        /// <Parameters> 
        ///     expression(ExpressionSyntax):
        ///     methodName(string):
        /// </Parameters>
        /// </METHOD>
        private bool IsEventHandlerReference(ExpressionSyntax expression, string methodName)
        {
            // Case 1: Simple identifier (e.g., "Button1_Click")
            if (expression is IdentifierNameSyntax identifier)
            {
                return identifier.Identifier.Text == methodName;
            }
            // Case 2: Member access (e.g., "this.Button1_Click")
            else if (expression is MemberAccessExpressionSyntax memberAccess)
            {
                return memberAccess.Name.Identifier.Text == methodName;
            }
            // If neither, it's not a match
            return false;
        }

        /// <METHOD name="doesWireEventHandler">
        /// <Purpose> 
        ///  Checks if the method wires any event handlers
        /// </Purpose>
        /// <Parameters> 
        ///     method(MethodDeclarationSyntax):
        ///     wiredEvents(List<string>):
        /// </Parameters>
        /// </METHOD>
        public bool doesWireEventHandler(MethodDeclarationSyntax method, out List<string> wiredEvents)
        {
            // Analyze the method's syntax tree to find += assignments
            wiredEvents = method.DescendantNodes()
                .OfType<AssignmentExpressionSyntax>()
                .Where(a => a.OperatorToken.Text == "+=" &&
                            (a.Left is MemberAccessExpressionSyntax || a.Left is IdentifierNameSyntax))
                .Select(a => a.Left.ToString())
                .ToList();

            // Return true if any events were wired, false otherwise
            return wiredEvents.Count > 0;
        }

        /// <METHOD name="attachTypesToParams">
        /// <Purpose> 
        ///  Attaches types to parameters in the XML comment
        /// </Purpose>
        /// <Parameters> 
        ///     paramString(string):
        ///     typeString(string):
        ///     indent(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     result:
        ///     result:
        ///     result:
        /// </Event_Raised>
        /// </METHOD>
        public string attachTypesToParams(string paramString, string typeString, string indent)
        {

            string[] splitString = paramString.Split(',');
            string[] splitTypes = typeString.Split('|');
            string result = "";
            for (int i = 0; i < splitString.Length; i++)
            {
                if (i == 0)
                {
                    result += "\n" + indent + "///     " + splitString[i] + "(" + splitTypes[i] + "):\n";
                }
                else
                {
                    result += indent + "///    " + splitString[i] + "(" + splitTypes[i] + "):\n";
                }


            }
            result += indent + "///";
            return result;
        }

        /// <METHOD name="removeNewLines">
        /// <Purpose> 
        ///  Removes new lines from the input string
        /// </Purpose>
        /// <Parameters> 
        ///     input(string):
        /// </Parameters>
        /// </METHOD>
        public string removeNewLines(string input)
        {
            return input.Replace("\n", "");
        }

        /// <METHOD name="convertExecptionToXML">
        /// <Purpose> 
        ///  Converts exception list to XML format
        /// </Purpose>
        /// <Parameters> 
        ///     exceptionList(List<string>):
        ///     indent(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     result:
        ///     result:
        ///     result:
        /// </Event_Raised>
        /// </METHOD>
        public string convertExecptionToXML(List<string> exceptionList, string indent)
        {

            string result = "";
            for (int i = 0; i < exceptionList.Count; i++)
            {
                if (i == 0)
                {
                    result += "\n" + indent + "///     " + exceptionList[i] + ":" + "\n";
                }
                else
                {
                    result += indent + "///     " + exceptionList[i] + ":" + "\n";
                }


            }
            result += indent + "///";
            return result;
        }

        /// <METHOD name="convertEventsToXML">
        /// <Purpose> 
        ///  Converts event list to XML format
        /// </Purpose>
        /// <Parameters> 
        ///     wires(List<string>):
        ///     indent(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     result:
        ///     result:
        ///     result:
        /// </Event_Raised>
        /// </METHOD>
        public string convertEventsToXML(List<string> wires, string indent)
        {
            string result = "";
            for (int i = 0; i < wires.Count; i++)
            {
                if (i == 0)
                {
                    result += "\n" + indent + "///" + "     " + wires[i] + ":\n";
                }
                else
                {
                    result += indent + "///" + "     " + wires[i] + ":\n";
                }

            }
            result += indent + "///";
            return result;

        }

        /// <METHOD name="findAndReplaceXML">
        /// <Purpose> 
        ///  Updates XML tags in the current document
        /// </Purpose>
        /// <Parameters> 
        ///     previous(string):
        ///     current(string):
        /// </Parameters>
        /// </METHOD>
        public void findAndReplaceXML(string previous, string current)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE2 dte = getDTE();
            if (dte?.ActiveDocument == null) return;
            Find find = dte.Find;
            find.Action = vsFindAction.vsFindActionReplaceAll;
            find.Target = vsFindTarget.vsFindTargetCurrentDocument;
            find.FindWhat = previous;
            find.ReplaceWith = current;
            find.MatchWholeWord = false;
            find.MatchCase = false;
            find.MatchInHiddenText = false;
            find.Execute();

        }

        /// <METHOD name="convertDynamicToXMLFormat">
        /// <Purpose> 
        ///  Converts dynamic text to XML format
        /// </Purpose>
        /// <Parameters> 
        ///     userText(string):
        ///     indent(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     result:
        ///     result:
        ///     result:
        /// </Event_Raised>
        /// </METHOD>
        public string convertDynamicToXMLFormat(string userText,string indent)
        {
            string filteredUserText = ReplaceTripleSlash(userText);
            string[] items = filteredUserText.Split('`');
            string result = "";
            for(int i=1; i<items.Length;i++)
            {
                if (i == 1)
                {
                    string updatedString = "\n" + indent + "///" + items[i] + "\n";
                    result += updatedString;
                }
                else
                {
                    string updatedString = indent + "///" + items[i] + "\n";
                    result += updatedString;
                }
            }
            result += indent + "///";
            return result;

        }

        /// <METHOD name="convertXMLStringToMultiLine">
        /// <Purpose> 
        ///  Convert XML string to multi-line format
        /// </Purpose>
        /// <Parameters> 
        ///     input(string):
        ///     indent(string):
        /// </Parameters>
        /// <Event_Raised> 
        ///     result:
        ///     result:
        ///     result:
        /// </Event_Raised>
        /// </METHOD>
        public string convertXMLStringToMultiLine(string input,string indent)
        {
            string updated = ReplaceTripleSlash(input);
            string[] items = updated.Split('`');
            string result = "";
            for (int i = 1; i < items.Length; i++)
            {
                if (i == 1)
                {
                    string updatedString = "\n" +indent+ "///"+ items[i] + "\n";
                    result += updatedString;
                }
                else
                {
                    string updatedString = "///"+items[i] + "\n";
                    result += updatedString;
                }
            }
            result += indent + "///";
            return result;
        }

        /// <METHOD name="ReplaceTripleSlash">
        /// <Purpose> 
        ///  Replaces triple slashes with a backtick
        /// </Purpose>
        /// <Parameters> 
        ///     input(string):
        /// </Parameters>
        /// </METHOD>
        public static string ReplaceTripleSlash(string input)
        {
            string updated = input.Replace("///", "`");
            return updated;
        }

        /// <METHOD name="btnUpdateAllTags_Click">
        /// <Purpose> 
        ///  Updates all tags in the current document
        /// </Purpose>
        /// <Parameters> 
        ///     sender(object):
        ///     e(RoutedEventArgs):
        /// </Parameters>
        /// </METHOD>
        private void btnUpdateAllTags_Click(object sender, RoutedEventArgs e)
        {
            SettingsManager sm = new SettingsManager();
            string[] toBecome = sm.names.Values.ToArray();
            findAndReplaceXML("/// <" + tagNames.tagCreatedBy, "/// <" + toBecome[0]);
            findAndReplaceXML("/// <" + tagNames.tagEventsHandled, "/// <" + toBecome[1]);
            findAndReplaceXML("/// <" + tagNames.tagEventRaised, "/// <" + toBecome[2]);
            findAndReplaceXML("/// <" + tagNames.tagExceptionThrown, "/// <" + toBecome[3]);
            findAndReplaceXML("/// <" + tagNames.tagExceptionCaught, "/// <" + toBecome[4]);
            findAndReplaceXML("/// <" + tagNames.tagPurpose, "/// <" + toBecome[5]);
            findAndReplaceXML("/// <" + tagNames.tagRevisionHistory, "/// <" + toBecome[6]);

            findAndReplaceXML("</" + tagNames.tagCreatedBy, "</" + toBecome[0]);
            findAndReplaceXML("</" + tagNames.tagEventsHandled, "</" + toBecome[1]);
            findAndReplaceXML("</" + tagNames.tagEventRaised, "</" + toBecome[2]);
            findAndReplaceXML("</" + tagNames.tagExceptionThrown, " </" + toBecome[3]);
            findAndReplaceXML("</" + tagNames.tagExceptionCaught, "</" + toBecome[4]);
            findAndReplaceXML("</" + tagNames.tagPurpose, "</" + toBecome[5]);
            findAndReplaceXML("</" + tagNames.tagRevisionHistory, "</" + toBecome[6]);
            tagNames.initalized = false;
        }
    }
}