using System;
using System.Diagnostics;
using EnvDTE;
using Microsoft.VisualStudio.CommandBars;

namespace RF.Ssms.Plugin.Common
{
    public class AddinInitializer
    {
        private const string C_Name = "RfSsmsPlugin";

        private SsmsAbstract _ssmsAbstract;
        private Window FToolwindowContainer = null;

        public void Initialize(SsmsAbstract ssmsAbstract)
        {
            _ssmsAbstract = ssmsAbstract;
            AddToolbarCommand(ssmsAbstract);
        }

        private void AddToolbarCommand(SsmsAbstract ssmsAbstract)
        {
            string toolbarName = C_Name + "Toolbar"; ;
            try
            {
                var toolbarControl = _ssmsAbstract.command_manager.create_toolbar_menu(
                "MenuBar",
                toolbarName,
                "RF ssms plugin command at Toolbar",
                0);

                _ssmsAbstract.command_manager.SuppressCommandBar(toolbarControl);

                try
                {
                    toolbarControl.Position = MsoBarPosition.msoBarTop;
                    toolbarControl.Left = 300;
                }
                catch
                {
                    Debug.WriteLine("Ignore the exception when set the " + nameof(toolbarControl));
                }

                AddToolbarItemCommand(toolbarName, toolbarControl);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Exception at AddToolbarCommand");
                Debug.WriteLine(ex.Message);
            }

            AddMenuCommand();
        }

        private void AddMenuCommand()
        {
            string manuName = C_Name + "Menu";

            try
            {
                _ssmsAbstract.command_manager.create_bar_command(
                            "Tools",
                            manuName,
                            "RF Ssms Plugin",
                            "RF ssms plugin command at Menu",
                            1,
                            null,
                            CreateOpenRfPluginCommandHandler(),
                            null,
                            true);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Exception at AddMenuCommand");
                Debug.WriteLine(ex.Message);
            }
        }

        private void AddToolbarItemCommand(string aToolbarName, CommandBar toptool)
        {
            if (toptool == null)
                return;

            foreach (CommandBarControl cbc in toptool.Controls)
            {
                cbc.Delete();
            }

            try
            {
                _ssmsAbstract.command_manager.create_bar_command(
                    aToolbarName,
                    "ItemCommandName",
                    "RF Ssms Plugin",
                    "RF ssms Plugin command in toolbar item",
                    1,
                    null,
                    CreateOpenRfPluginCommandHandler(),
                    null,
                    false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Exception at AddToolbarItemCommand");
                Debug.WriteLine(ex.Message);
            }
        }

        private menu_command_exec_handler CreateOpenRfPluginCommandHandler()
        {
            menu_command_exec_handler exec_handler = delegate (vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
            {
                if (FToolwindowContainer != null)
                {
                    FToolwindowContainer.Activate();
                }
                else
                {
                    try
                    {
                        var wins2Obj = (EnvDTE80.Windows2)_ssmsAbstract.DTE2.Windows;

                        object myControl = null;

                        FToolwindowContainer = wins2Obj.CreateToolWindow2(
                                 _ssmsAbstract.addin,
                                 System.Reflection.Assembly.GetExecutingAssembly().Location,
                                 typeof(SQL2014.WinformHost).FullName, // Should be use Windows Forms control
                                 "RF Wpf User Control",
                                 "{F198702E-02A2-42D7-AE75-E4ECFB0E5686}",
                                 ref myControl
                        );
                        //var _wpfToolWindow = myControl as WinformHost;

                        FToolwindowContainer.IsFloating = false;
                        FToolwindowContainer.Visible = true;
                        FToolwindowContainer.Linkable = false;
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Exception when Open:" + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace);
                    }
                }
            };

            return exec_handler;
        }
    }
}