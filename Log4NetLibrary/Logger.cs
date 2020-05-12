using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace Log4NetLibrary
{
    public static class Logger
    {
        private static ILog _log;

        public static void EnsureLogger()
        {
            if (_log != null) return;

            var assembly = Assembly.GetEntryAssembly();
            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            var configFile = GetConfigFile();

            // Configure Log4Net
            XmlConfigurator.Configure(logRepository, configFile);
            _log = LogManager.GetLogger(assembly, assembly.ManifestModule.Name.Replace(".dll", "").Replace(".", " "));
        }

        public static void Debug(object message)
        {
            EnsureLogger();
            _log.Debug(message);
        }

        public static void Error(object message)
        {
            EnsureLogger();
            _log.Error(message);
        }

        public static void Info(object message)
        {
            EnsureLogger();
            _log.Info(message);
        }

        public static void Warn(object message)
        {
            EnsureLogger();
            _log.Warn(message);
        }


        private static FileInfo GetConfigFile()
        {
            FileInfo configFile = null;

            // Search config file
            var configFileNames = new[] { "Config/log4net.config", "log4net.config" };

            foreach (var configFileName in configFileNames)
            {
                configFile = new FileInfo(configFileName);

                if (configFile.Exists) break;
            }

            //  if (configFile == null || !configFile.Exists) throw new NullReferenceException("Log4net config file not found.");

            return configFile;
        }

        public static void EnterScope()
        {
            //get calling method
            StackTrace stackTrace = new StackTrace();
            MethodBase methodBase = stackTrace.GetFrame(1).GetMethod();

            string message = string.Format("Enter Scope : {0}", methodBase.Name);

            Logger.Debug(message);
        }

        public static void ExitScope()
        {
            //get calling method
            StackTrace stackTrace = new StackTrace();
            MethodBase methodBase = stackTrace.GetFrame(1).GetMethod();

            string message = string.Format("Exit Scope : {0}", methodBase.Name);

            Logger.Debug(message);
        }
    }

}
