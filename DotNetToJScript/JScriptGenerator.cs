//    This file is part of DotNetToJScript - A tool to generate a 
//    JScript which bootstraps an arbitrary v2.NET Assembly and class.
//    Copyright (C) James Forshaw 2017
//
//    DotNetToJScript is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    DotNetToJScript is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with DotNetToJScript.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Text;

namespace DotNetToJScript
{
    class JScriptGenerator : IScriptGenerator
    {
        static string debug_enabled = @"function debug(s) { WScript.Echo(s); }" + Environment.NewLine;
        static string debug_disabled = @"function debug(s) { }" + Environment.NewLine;

        static string GetScriptHeader(RuntimeVersion version, bool enable_debug)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("function setversion() {");
            switch (version)
            {
                case RuntimeVersion.Auto:
                    builder.AppendLine(Properties.Resources.jscript_auto_version_script);
                    break;
                case RuntimeVersion.v2:
                    builder.AppendLine("new ActiveXObject('WScript.Shell').Environment('Process')('COMPLUS_Version') = 'v2.0.50727';");
                    break;
                case RuntimeVersion.v4:
                    builder.AppendLine("new ActiveXObject('WScript.Shell').Environment('Process')('COMPLUS_Version') = 'v4.0.30319';");
                    break;
            }
            builder.AppendLine("}");
            builder.AppendLine("function debug(s) {");
            if (enable_debug)
            {
                builder.AppendLine("WScript.Echo(s);");
            }
            builder.AppendLine("}");
            return builder.ToString();
        }

        public string ScriptName
        {
            get
            {
                return "JScript";
            }
        }

        public bool SupportsScriptlet
        {
            get
            {
                return true;
            }
        }

        public string GenerateScript(byte[] serialized_object, string entry_class_name, string additional_script, RuntimeVersion version, bool enable_debug)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < serialized_object.Length; ++i)
            {
                builder.Append(serialized_object[i]);
                if (i < serialized_object.Length - 1)
                {
                    builder.Append(",");
                }
                if (i > 0 && (i % 32) == 0)
                {
                    builder.AppendLine();
                }
            }
            
            return GetScriptHeader(version, enable_debug) 
                + Properties.Resources.jscript_template.Replace("%SERIALIZED%", builder.ToString()).Replace("%CLASS%", entry_class_name).Replace("%ADDEDSCRIPT%", additional_script);
        }
    }
}
