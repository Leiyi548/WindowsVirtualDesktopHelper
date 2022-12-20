﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsVirtualDesktopHelper.VirtualDesktopAPI {
    public class Loader {

        public const string VirtualDesktopWin11_22H2 = "VirtualDesktopWin11_22H2";
        public const string VirtualDesktopWin11_21H2 = "VirtualDesktopWin11_21H2";
        public const string VirtualDesktopWin10 = "VirtualDesktopWin10";

        public static string GetImplementationForOS() {
            // We need to load the correct API for correct windows version...
            // See https://www.anoopcnair.com/windows-11-version-numbers-build-numbers-major/ for versions
            int currentBuild = 0;
            try {
                currentBuild = Util.OS.GetWindowsBuildVersion();
            } catch (Exception e) {
                throw new Exception("LoadVDAPI: could not determine Windows version: " + e.Message, e);
            }
            Util.Logging.WriteLine("GetImplementationForOS: Windows Build Version: " + currentBuild);
            if (currentBuild >= 22621) {
                Util.Logging.WriteLine("GetImplementationForOS: Detected Windows 11 22H2 due to build >= 22621");
                return VirtualDesktopWin11_22H2;
            } else if (currentBuild >= 22000) {
                Util.Logging.WriteLine("GetImplementationForOS: Detected Windows 11 21H1 due to build >= 22000");
                return VirtualDesktopWin11_21H2;
            } else if (currentBuild >= 21996) {
                Util.Logging.WriteLine("GetImplementationForOS: Detected Windows 11 due to build >= 21996 (beta version of Windows 11)");
                return VirtualDesktopWin11_21H2;
            } else {
                Util.Logging.WriteLine("GetImplementationForOS: Fallback to Windows 10 (fallback)");
                return VirtualDesktopWin10;
            }
        }

        public static IVirtualDesktopManager LoadImplementationWithFallback(string name) {
            var implementationsToTry = new List<string>();
            implementationsToTry.Add(name);
            if (!implementationsToTry.Contains(VirtualDesktopWin11_22H2)) implementationsToTry.Add(VirtualDesktopWin11_22H2);
            if (!implementationsToTry.Contains(VirtualDesktopWin11_21H2)) implementationsToTry.Add(VirtualDesktopWin11_21H2);
            if (!implementationsToTry.Contains(VirtualDesktopWin10)) implementationsToTry.Add(VirtualDesktopWin10);

            foreach(var implementationName in implementationsToTry) {
                Util.Logging.WriteLine("LoadImplementationWithFallback: trying to load implementation " + implementationName);
                try {
                    var impl = LoadImplementation(name);
                    impl.Current(); // test for success
                    Util.Logging.WriteLine("LoadImplementationWithFallback: success!");
                    return impl;
                }catch(Exception e) {
                    Util.Logging.WriteLine("LoadImplementationWithFallback: failed to load " + implementationName);
                }
            }
            throw new Exception("LoadImplementationWithFallback: no implementation loaded successfully, tried: "+string.Join(",",implementationsToTry));
        }

        public static IVirtualDesktopManager LoadImplementation(string name) {
            Util.Logging.WriteLine("LoadImplementation: Loading VDImplementation: " + name + "...");
            IVirtualDesktopManager impl = null;
            if (name == VirtualDesktopWin11_22H2) {
                try {
                    impl = new VirtualDesktopAPI.Implementation.VirtualDesktopWin11_22H2();
                } catch (Exception e) {
                    throw new Exception("LoadImplementation: could not load VirtualDesktop API implementation "+name+": " + e.Message, e);
                }
            } else if (name == VirtualDesktopWin11_21H2) {
                try {
                    impl = new VirtualDesktopAPI.Implementation.VirtualDesktopWin11_21H2();
                } catch (Exception e) {
                    throw new Exception("LoadImplementation: could not load VirtualDesktop API implementation " + name + ": " + e.Message, e);
                }
            } else if(name == VirtualDesktopWin10) {
                try {
                    impl = new VirtualDesktopAPI.Implementation.VirtualDesktopWin10();
                } catch (Exception e) {
                    throw new Exception("LoadImplementation: could not load VirtualDesktop API implementation " + name + ": " + e.Message, e);
                }
            } else {
                throw new Exception("LoadImplementation: could not load VirtualDesktop API implementation " + name + ": " + "Unknown implementation");
            }
            return impl;
        }
    }
}
