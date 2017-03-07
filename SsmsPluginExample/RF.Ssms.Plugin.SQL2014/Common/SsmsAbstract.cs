using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using EnvDTE;
using EnvDTE80;
using BARS = Microsoft.VisualStudio.CommandBars;
using VSI = Microsoft.SqlServer.Management.UI.VSIntegration;

namespace RF.Ssms.Plugin.Common
{
    public delegate void menu_command_querystatus_handler(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText);

    public delegate void menu_command_exec_handler(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled);

    public class query_window
    {
        private Document _doc;
        private TextDocument _tdoc;

        public query_window set_doc(Document doc)
        {
            _doc = doc;
            _tdoc = (TextDocument)_doc.Object("TextDocument");
            return this;
        }

        public string text
        {
            get
            {
                EditPoint p = _tdoc.StartPoint.CreateEditPoint();
                return p.GetText(_tdoc.EndPoint);
            }
            set
            {
                EditPoint p1 = _tdoc.StartPoint.CreateEditPoint();
                EditPoint p2 = _tdoc.EndPoint.CreateEditPoint();
                p1.Delete(p2);
                p1.Insert(value);
            }
        }

        public string current_query
        {
            get
            {
                if (_tdoc.Selection != null && _tdoc.Selection.Text != "")
                {
                    return _tdoc.Selection.Text;
                }
                else
                {
                    return this.text;
                }
            }
        }

        public string current_select_query
        {
            get
            {
                if (_tdoc.Selection != null && _tdoc.Selection.Text != "")
                {
                    return _tdoc.Selection.Text;
                }

                return String.Empty;
            }
        }

        public string name
        {
            get
            {
                return _doc.Name;
            }
        }

        public string file_path
        {
            get
            {
                return _doc.FullName;
            }
        }

        public void insert_in_current(string value)
        {
            _tdoc.Selection.ActivePoint.CreateEditPoint().Insert(value);
            EditPoint p2 = _tdoc.EndPoint.CreateEditPoint();
        }
    }

    public class window_manager
    {
        private SsmsAbstract _ssms;
        private Windows2 _SSMS_window_collection;
        private query_window _working_query_window;
        private System.Collections.Generic.List<Window> _addin_window_collection;

        public query_window query_windows(int index)
        {
            return _ssms.DTE2.Documents.Item(index) == null ? null : _working_query_window.set_doc(_ssms.DTE2.ActiveDocument);
        }

        public query_window active_query_window
        {
            get
            {
                return _ssms.DTE2.ActiveDocument == null ? null : _working_query_window.set_doc(_ssms.DTE2.ActiveDocument);
            }
        }

        public query_window create_query_window()
        {
            EnvDTE.TextDocument doc;

            VSI.ServiceCache.ScriptFactory.CreateNewBlankScript(VSI.Editors.ScriptType.Sql);
            doc = (EnvDTE.TextDocument)VSI.ServiceCache.ExtensibilityModel.Application.ActiveDocument.Object(null);
            doc.EndPoint.CreateEditPoint().Insert("");

            query_window queryWindow = new query_window();
            queryWindow.set_doc(VSI.ServiceCache.ExtensibilityModel.ActiveDocument);

            return queryWindow;
        }

        public window_manager(SsmsAbstract ssms)
        {
            _ssms = ssms;
            _working_query_window = new query_window();
            _SSMS_window_collection = (Windows2)_ssms.DTE2.Windows;
            _addin_window_collection = new System.Collections.Generic.List<Window>();
        }
    }

    public class command_manager
    {
        private struct menu_command_handlers
        {
            public menu_command_querystatus_handler querystatus_handler;
            public menu_command_exec_handler exec_handler;
        }

        private System.Collections.Generic.List<BARS.CommandBar> _addin_command_bars;
        private Commands2 _SSMS_commands_collection;
        private System.Collections.Generic.Dictionary<string, menu_command_handlers> _addin_menu_commands_dictonary;
        private SsmsAbstract _ssms;

        public command_manager(SsmsAbstract ssms)
        {
            _SSMS_commands_collection = (Commands2)VSI.ServiceCache.ExtensibilityModel.Commands;
            _addin_menu_commands_dictonary = new System.Collections.Generic.Dictionary<string, menu_command_handlers>();
            _ssms = ssms;
            _addin_command_bars = new System.Collections.Generic.List<BARS.CommandBar>();
        }

        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (_addin_menu_commands_dictonary.ContainsKey(commandName))
                {
                    _addin_menu_commands_dictonary[commandName].querystatus_handler(neededText, ref status, ref commandText);
                }
            }
        }

        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (_addin_menu_commands_dictonary.ContainsKey(commandName))
            {
                _addin_menu_commands_dictonary[commandName].exec_handler(executeOption, ref varIn, ref varOut, ref handled);
            }
        }

        private void _default_querystatus_handler(vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
        }

        private void _default_exec_handler(vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = true;
        }

        public void SuppressCommandBar(BARS.CommandBar bar)
        {
            if (_addin_command_bars.Contains(bar))
                _addin_command_bars.Remove(bar);
        }

        private void RemoveCommandIfAlreadyExits(string command_name)
        {
            try
            {
                Command cmd = _SSMS_commands_collection.Item(_ssms.addin.ProgID + "." + command_name);
                cmd.Delete();
            }
            catch { }
        }

        public BARS.CommandBar create_toolbar_menu(string host_menu_bar_name, string new_bar_name, string tooltip_text, int position)
        {
            BARS.CommandBar new_bar = null;
            // check existing
            try
            {
                new_bar = (BARS.CommandBar)((BARS.CommandBars)VSI.ServiceCache.ExtensibilityModel.CommandBars)[new_bar_name];
            }
            catch
            {
            }
            if (new_bar == null)
            {
                try
                {
                    BARS.CommandBar host_bar = ((BARS.CommandBars)VSI.ServiceCache.ExtensibilityModel.CommandBars)[host_menu_bar_name];
                    position = (position == 0 ? host_bar.Controls.Count + 1 : position);
                    new_bar = (BARS.CommandBar)_SSMS_commands_collection.AddCommandBar(new_bar_name, vsCommandBarType.vsCommandBarTypeToolbar, host_bar, position);
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.Print(e.Message);
                }
            }

            if (new_bar != null)
            {
                _addin_command_bars.Add(new_bar);
                new_bar.Visible = true;
            }

            return new_bar;
        }

        public void create_bar_command(
           string host_menu_bar_name,
           string command_name,
           string item_text,
           string tooltip_text,
           int position,
           menu_command_querystatus_handler querystatus_handler,
           menu_command_exec_handler exec_handler,
           Bitmap picture,
           bool beginGroup
        )
        {
            try
            {
                object[] contextGUIDS = new object[] { };
                RemoveCommandIfAlreadyExits(command_name);

                vsCommandStyle commandStyle = (picture == null) ? vsCommandStyle.vsCommandStyleText : vsCommandStyle.vsCommandStylePictAndText;

                BARS.CommandBar host_bar = (BARS.CommandBar)((BARS.CommandBars)VSI.ServiceCache.ExtensibilityModel.CommandBars)[host_menu_bar_name];
                Command new_command = _SSMS_commands_collection.AddNamedCommand2(_ssms.addin, command_name, item_text, tooltip_text, true, 0, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)commandStyle, vsCommandControlType.vsCommandControlTypeButton);
                if (new_command != null)
                {
                    position = (position == 0 ? host_bar.Controls.Count + 1 : position);
                    BARS.CommandBarButton control = (BARS.CommandBarButton)new_command.AddControl(host_bar, position);
                    control.BeginGroup = beginGroup;
                    control.TooltipText = tooltip_text.Substring(0, tooltip_text.Length > 255 ? 255 : tooltip_text.Length);
                    menu_command_handlers handlers = new menu_command_handlers();
                    handlers.querystatus_handler = querystatus_handler == null ? this._default_querystatus_handler : querystatus_handler;
                    handlers.exec_handler = exec_handler == null ? this._default_exec_handler : exec_handler;
                    this._addin_menu_commands_dictonary.Add(_ssms.addin.ProgID + "." + command_name, handlers);
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception at create_bar_command");
                System.Diagnostics.Debug.Print(e.Message);
            }
        }

        public void cleanup()
        {
            var array = _addin_command_bars.ToArray();

            for (int i = 0; i < array.Length; i++)
            {
                array[i].Delete();
                _addin_command_bars.Remove(array[i]);
            }

            foreach (Command c in VSI.ServiceCache.ExtensibilityModel.Commands)
            {
                if (_addin_menu_commands_dictonary.ContainsKey(c.Name))
                {
                    string name = c.Name;
                    c.Delete();
                    _addin_menu_commands_dictonary.Remove(name);
                }
            }
        }

        public void CustomCleanup()
        {
            foreach (Command command in VSI.ServiceCache.ExtensibilityModel.Commands)
            {
                DeleteMenuCommandHandlers(command);
            }

            foreach (BARS.CommandBar commandBar in _addin_command_bars)
            {
                DeleteCommandBarControl(commandBar);
            }
        }

        private void DeleteMenuCommandHandlers(Command aCommand)
        {
            if (_addin_menu_commands_dictonary.ContainsKey(aCommand.Name))
            {
                aCommand.Delete();
                _addin_menu_commands_dictonary.Remove(aCommand.Name);
            }
        }

        private void DeleteCommandBarControl(BARS.CommandBar aCommandBar)
        {
            try
            {
                List<object> controls = new List<object>();
                foreach (var control in aCommandBar.Controls)
                {
                    controls.Add(control);
                }

                foreach (var control in controls)
                {
                    if (control is BARS.CommandBarPopup)
                    {
                        (control as BARS.CommandBarPopup).Delete();
                        continue;
                    }

                    if (control is BARS.CommandBarControl)
                    {
                        (control as BARS.CommandBarControl).Delete();
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }
    }

    public class SsmsAbstract
    {
        private EnvDTE80.DTE2 _DTE2;
        private EnvDTE.DTE _DTE;
        private AddIn _addin;
        private command_manager _command_manager;
        private window_manager _window_manager;

        public SsmsAbstract(AddIn addin_instance, DTE2 applicationObject)
        {
            _addin = addin_instance;
            _DTE2 = applicationObject;
            _DTE = (DTE)applicationObject;
            _command_manager = new command_manager(this);
            _window_manager = new window_manager(this);
        }

        public EnvDTE80.DTE2 DTE2
        {
            get
            {
                return _DTE2;
            }
        }

        public EnvDTE.DTE DTE
        {
            get
            {
                return _DTE;
            }
        }

        public AddIn addin
        {
            get
            {
                return _addin;
            }
        }

        public command_manager command_manager
        {
            get
            {
                return _command_manager;
            }
        }

        public window_manager window_manager
        {
            get
            {
                return _window_manager;
            }
        }

        public void OnDisconnection()
        {
            _command_manager.cleanup();
        }
    }
}