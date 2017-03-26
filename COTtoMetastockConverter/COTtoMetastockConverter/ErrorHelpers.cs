using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;

namespace COTtoMetastockConverter
{
    public static class ErrorHelpers
    {
        private static List<string> _errorsList = new List<string>();

        //raises an immediate exception
        public static void immediateEx(string msg)
        {
            string err = String.Format(msg);
            throw new Exception(err);
        }
        //saves the error messages in the _errorsList. An exception can be thrown later
        public static void deferredEx(string msg)
        {
            _errorsList.Add(msg);
        }
        public static List<string> errorMessages()
        {
            if (_errorsList == null || _errorsList.Count == 0) return new List<string>();
            else return _errorsList;
        }
    }
}
