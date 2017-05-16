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

        public string GenerateScript(byte[] serialized_object, string entry_class_name, string additional_script, RuntimeVersion version, bool enable_debug)
        {
            if (version != RuntimeVersion.None)
            {
                throw new ArgumentException("VBA generator does not support version detection");
            }

            if (enable_debug)
            {
                throw new ArgumentException("VBA generator does not support debug output");
            }

            string hex_encoded = BitConverter.ToString(serialized_object).Replace("-", "");
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < hex_encoded.Length; i++)
            {
                if (i == 0)
                {
                    builder.Append("    serialized_obj = \"");
                }
                else if (i % 100 == 0)
                {
                    builder.Append("\"");
                    builder.AppendLine();
                    builder.Append("    serialized_obj = serialized_obj & \"");
                }
                builder.Append(hex_encoded[i]);
            }
            builder.Append("\"");

            return Properties.Resources.vba_template.Replace("%SERIALIZED%", builder.ToString()).Replace("%CLASS%", entry_class_name).Replace("%ADDEDSCRIPT%", additional_script);
        }
    }
}
