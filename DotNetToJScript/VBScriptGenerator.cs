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
    static class VBShared
    {
        public static string GetScriptHeader(RuntimeVersion version, bool enable_debug, string script_name)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Sub DebugPrint(s)");
            if (enable_debug)
            {
                if (script_name == "VBScript") builder.AppendLine("Wscript.Echo s");
                else if (script_name == "VBA") builder.AppendLine("Debug.Print s");
            }
            builder.AppendLine("End Sub");
            builder.AppendLine();

            builder.AppendLine("Sub SetVersion");
            if (version != RuntimeVersion.None)
            {
                builder.AppendLine("Dim shell");
                builder.AppendLine("Set shell = CreateObject(\"WScript.Shell\")");
                switch (version)
                {
                    case RuntimeVersion.v2:
                        builder.AppendLine("shell.Environment(\"Process\").Item(\"COMPLUS_Version\") = \"v2.0.50727\"");
                        break;
                    case RuntimeVersion.v4:
                        builder.AppendLine("shell.Environment(\"Process\").Item(\"COMPLUS_Version\") = \"v4.0.30319\"");
                        break;
                    case RuntimeVersion.Auto:
                        builder.AppendLine(Properties.Resources.vb_multi_auto_version_script);
                        break;
                }
            }
            builder.AppendLine("End Sub");
            builder.AppendLine();
            return builder.ToString();
        }
    }
    class VBScriptGenerator : IScriptGenerator
    {
        public string ScriptName
        {
            get
            {
                return "VBScript";
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
            string[] lines = JScriptGenerator.BinToBase64Lines(serialized_object);

            return VBShared.GetScriptHeader(version, enable_debug, ScriptName) + Properties.Resources.vbs_template.Replace("%SERIALIZED%", 
                String.Join(Environment.NewLine + "s = s & ", lines)).Replace("%CLASS%", entry_class_name).Replace("%ADDEDSCRIPT%", additional_script);
        }
    }
}
