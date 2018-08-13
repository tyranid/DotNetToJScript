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
    class VBAGenerator : IScriptGenerator
    {
        public string ScriptName
        {
            get
            {
                return "VBA";
            }
        }

        public bool SupportsScriptlet
        {
            get
            {
                return false;
            }
        }

        public static string GetScriptHeader(Boolean debug)
        {
            StringBuilder builder = new StringBuilder();

            builder.AppendLine("Sub DebugPrint(s)");
            if (debug) builder.AppendLine("    Debug.Print s");
            builder.AppendLine("End Sub");
            builder.AppendLine();

            return builder.ToString();
        }

        public static string GetManifest(RuntimeVersion version)
        {
            StringBuilder builder = new StringBuilder();

            string runtimeVersion = (version != RuntimeVersion.v2) ? "v4.0.30319" : "v2.0.50727";
            string mscorlibVersion = (version != RuntimeVersion.v2) ? "4.0.0.0" : "2.0.0.0";

            string template = Properties.Resources.manifest_template.Replace(
                                                                    "%RUNTIMEVERSION%",
                                                                    runtimeVersion
                                                                ).Replace(
                                                                    "%MSCORLIBVERSION%",
                                                                    mscorlibVersion
                                                                );

            for (int i = 0; i < template.Length; i++)
            {
                if (i == 0)
                {
                    builder.Append("manifest = \"");
                }
                else if (i % 300 == 0)
                {
                    builder.Append("\"");
                    builder.AppendLine();
                    builder.Append("        manifest = manifest & \"");
                }
                builder.Append(template[i]);
                if (template[i] == '"') builder.Append('"');
            }
            builder.Append("\"");

            return builder.ToString();
        }

        public string GenerateScript(byte[] serialized_object, string entry_class_name, string additional_script, RuntimeVersion version, bool enable_debug)
        {
            string hex_encoded = BitConverter.ToString(serialized_object).Replace("-", "");
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < hex_encoded.Length; i++)
            {
                if (i == 0)
                {
                    builder.Append("s = \"");
                }
                else if (i % 300 == 0)
                {
                    builder.Append("\"");
                    builder.AppendLine();
                    builder.Append("    s = s & \"");
                }
                builder.Append(hex_encoded[i]);
            }
            builder.Append("\"");

            return GetScriptHeader(enable_debug) +
                Properties.Resources.vba_template.Replace(
                                                "%SERIALIZED%",
                                                builder.ToString()
                                            ).Replace(
                                                "%CLASS%",
                                                entry_class_name
                                            ).Replace(
                                                "%MANIFEST%",
                                                GetManifest(version)
                                            ).Replace(
                                                "%ADDEDSCRIPT%",
                                                additional_script
                                            );
        }
    }
}
