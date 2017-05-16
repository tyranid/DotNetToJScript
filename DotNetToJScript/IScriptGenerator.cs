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

namespace DotNetToJScript
{
    enum RuntimeVersion
    {
        None,
        v2,
        v4,
        Auto,
    }

    interface IScriptGenerator
    {
        string GenerateScript(byte[] serialized_object,
                              string entry_class_name,
                              string additional_script, 
                              RuntimeVersion version,
                              bool enable_debug);
        bool SupportsScriptlet { get; }
        string ScriptName { get; }
    }
}
