using System;
using System.IO;
using System.Reflection;

namespace RF.Ssms.Plugin.Common
{
    public static class AssemblyLoader
    {
        public static void LinkAssemblyResolveEventToReloadAssembly()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return LoadOwnAssemblyWhenPluginLoadFailed(args.Name);
        }

        private static Assembly LoadOwnAssemblyWhenPluginLoadFailed(string aAssemblyFullName)
        {
            try
            {
                var currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(currentDir))
                {
                    var assembly = new AssemblyName(aAssemblyFullName);
                    var assemblyDllName = assembly.Name + ".dll";

                    var fullPath = Path.Combine(currentDir, assemblyDllName);
                    if (File.Exists(fullPath))
                    {
                        return Assembly.LoadFile(fullPath);
                    }
                }
            }
            catch
            {
                // Ignore the exception.
            }

            return null;
        }
    }
}