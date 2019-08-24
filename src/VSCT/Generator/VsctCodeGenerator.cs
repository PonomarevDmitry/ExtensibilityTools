using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using EnvDTE;

namespace MadsKristensen.ExtensibilityTools.VSCT.Generator
{
    /// <summary>
    /// This is the dedicated code generator class that processes VSCT files and outputs their CommandIDs/GUIDs to be used in code.
    /// When setting the 'Custom Tool' property of a C#, VB, or J# project item to "VsctGenerator", 
    /// the GenerateCode function will get called and will return the contents of the generated file to the project system.
    /// </summary>
    [Guid("a6a34300-fa6b-4f86-a8ba-e1fea8d24922")]
    public sealed class VsctCodeGenerator : BaseCodeGenerator
    {
        #region Public Names

        /// <summary>
        /// Name of this generator.
        /// </summary>
        public const string GeneratorName = "VsctGenerator";

        /// <summary>
        /// Description of this generator.
        /// </summary>
        public const string GeneratorDescription = "Generates .NET source code for given VS IDE GUI definitions.";

        #endregion

        #region Comments

        private const string DefaultGuidListClassName = "PackageGuids";
        private const string DefaultPkgCmdIDListClassName = "PackageIds";

        private const string ClassGuideListComment = "Helper class that exposes all GUIDs used across VS Package.";
        private const string ClassPkgCmdIDListComments = "Helper class that encapsulates all CommandIDs uses across VS Package.";

        #endregion

        protected override string GenerateStringCode(string inputFileContent)
        {
            // get parameters passed as 'FileNamespace' inside properties of the file generator:
            InterpreteArguments(
                string.IsNullOrEmpty(FileNamespace) ? null : FileNamespace.Split(';')
                , out string globalNamespaceName
                , out string guidListClassName
                , out string cmdIdListClassName
                , out string supporterPostfix
                , out bool isPublic
            );

            string language = GetProject().CodeModel.Language;

            var codeProvider = GetCodeProvider(language);

            bool isCSharp = string.Equals(codeProvider.FileExtension, "cs", StringComparison.InvariantCultureIgnoreCase);

            // create support CodeDOM classes:
            var globalNamespace = new System.CodeDom.CodeNamespace(globalNamespaceName);

            var classGuideList = CreateClass(guidListClassName, ClassGuideListComment, isPublic, true, !isCSharp);

            var classPkgCmdIDList = CreateClass(cmdIdListClassName, ClassPkgCmdIDListComments, isPublic, true, !isCSharp);

            // retrieve the list GUIDs and IDs defined inside VSCT file:
            var guids = Parse(inputFileContent);

            // generate members describing GUIDs:
            if (guids != null)
            {
                var delayedMembers = new List<CodeTypeMember>();

                foreach (var symbol in guids)
                {
                    string nameString;
                    string nameGuid;

                    // for each GUID generate one string and one GUID field with similar names:
                    if (symbol.Name != null && symbol.Name.EndsWith(supporterPostfix, StringComparison.OrdinalIgnoreCase))
                    {
                        nameString = symbol.Name;
                        nameGuid = symbol.Name.Substring(0, symbol.Name.Length - supporterPostfix.Length);
                    }
                    else
                    {
                        nameString = symbol.Name + supporterPostfix;
                        nameGuid = symbol.Name;
                    }

                    classGuideList.Members.Add(CreateConstField(typeof(string), nameString, symbol.Value, true));

                    delayedMembers.Add(CreateStaticField(typeof(Guid), nameGuid, nameString, true));

                    if (symbol.Ids.Any())
                    {
                        if (classPkgCmdIDList.Members.Count > 0)
                        {
                            classPkgCmdIDList.Members.Add(new CodeSnippetTypeMember(string.Empty));
                        }

                        var classPkgCmdIDNested = CreateClass(nameGuid, string.Empty, isPublic, true, !isCSharp);

                        classPkgCmdIDList.Members.Add(classPkgCmdIDNested);

                        foreach (var id in symbol.Ids)
                        {
                            if (classPkgCmdIDNested.Members.Count > 0)
                            {
                                classPkgCmdIDNested.Members.Add(new CodeSnippetTypeMember(string.Empty));
                            }

                            classPkgCmdIDNested.Members.Add(CreateConstField(typeof(int), id.Item1, ToHex(id.Item2, language), false));
                        }
                    }
                }

                classGuideList.Members.Add(new CodeSnippetTypeMember(string.Empty));

                foreach (var member in delayedMembers)
                {
                    classGuideList.Members.Add(member);
                }
            }

            globalNamespace.Comments.Add(new CodeCommentStatement("------------------------------------------------------------------------------"));
            globalNamespace.Comments.Add(new CodeCommentStatement("<auto-generated>"));
            globalNamespace.Comments.Add(new CodeCommentStatement($"    This file was generated by {Vsix.Name} v{Vsix.Version}"));
            globalNamespace.Comments.Add(new CodeCommentStatement("</auto-generated>"));
            globalNamespace.Comments.Add(new CodeCommentStatement("------------------------------------------------------------------------------"));

            globalNamespace.Types.Add(classGuideList);
            globalNamespace.Types.Add(classPkgCmdIDList);

            // generate source code:
            return GenerateFromNamespace(codeProvider, globalNamespace, false, isCSharp);
        }

        private void InterpreteArguments(string[] args, out string globalNamespaceName, out string guidClassName, out string cmdIdListClassName, out string supporterPostfix, out bool isPublic)
        {
            globalNamespaceName = GetProject().Properties.Item("DefaultNamespace").Value as string;
            guidClassName = DefaultGuidListClassName;
            cmdIdListClassName = DefaultPkgCmdIDListClassName;
            supporterPostfix = "String";
            isPublic = false;

            if (args != null && args.Length != 0)
            {
                if (!string.IsNullOrEmpty(args[0]))
                    globalNamespaceName = args[0];

                if (!(args.Length < 2 || string.IsNullOrEmpty(args[1])))
                    guidClassName = args[1];

                if (!(args.Length < 3 || string.IsNullOrEmpty(args[2])))
                    cmdIdListClassName = args[2];

                if (!(args.Length < 4 || string.IsNullOrEmpty(args[3])))
                    supporterPostfix = args[3];

                if (args.Length >= 5 && !string.IsNullOrEmpty(args[4]) && string.Compare(args[4], "public", StringComparison.OrdinalIgnoreCase) == 0)
                    isPublic = true;
            }
        }

        #region Parsing

        private class PackageGuid
        {
            public string Name { get; private set; }

            public string Value { get; private set; }

            public List<Tuple<string, string>> Ids { get; private set; }

            public PackageGuid(string name, string value)
            {
                this.Name = name;
                this.Value = value;

                this.Ids = new List<Tuple<string, string>>();
            }
        }

        /// <summary>
        /// Extract GUIDs and IDs descriptions from given XML content.
        /// </summary>
        private static List<PackageGuid> Parse(string vsctContentFile)
        {
            List<PackageGuid> result = new List<PackageGuid>();

            var xml = new XmlDocument();

            XmlElement symbols = null;

            try
            {
                xml.LoadXml(vsctContentFile);

                // having XML loaded go through and find:
                // CommandTable / Symbols / GuidSymbol* / IDSymbol*
                if (xml.DocumentElement != null && xml.DocumentElement.Name == "CommandTable")
                    symbols = xml.DocumentElement["Symbols"];
            }
            catch
            {
                return result;
            }

            if (symbols != null)
            {
                var guidSymbols = symbols.GetElementsByTagName("GuidSymbol");

                foreach (XmlElement symbol in guidSymbols)
                {
                    try
                    {
                        // go through all GuidSymbol elements...
                        var value = symbol.Attributes["value"].Value;
                        var name = symbol.Attributes["name"].Value;

                        // preprocess value to remove the brackets:
                        try
                        {
                            value = new Guid(value).ToString("D");
                        }
                        catch
                        {
                            value = "-invalid-";
                        }

                        var guid = new PackageGuid(name, value);

                        result.Add(guid);

                        var idSymbols = symbol.GetElementsByTagName("IDSymbol");

                        foreach (XmlElement i in idSymbols)
                        {
                            try
                            {
                                // go through all IDSymbol elements...
                                guid.Ids.Add(Tuple.Create(i.Attributes["name"].Value, i.Attributes["value"].Value));
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return result;
        }

        #endregion

        #region Code Definition

        /// <summary>
        /// Creates new static/partial class definition.
        /// </summary>
        private static CodeTypeDeclaration CreateClass(string name, string comment, bool isPublic, bool isPartial, bool isSealed)
        {
            var item = new CodeTypeDeclaration(name);

            if (!string.IsNullOrEmpty(comment))
            {
                item.Comments.Add(new CodeCommentStatement("<summary>", true));
                item.Comments.Add(new CodeCommentStatement(comment, true));
                item.Comments.Add(new CodeCommentStatement("</summary>", true));
            }

            item.IsClass = true;

            // HINT: Sealed + Abstract => static class definition
            if (isPublic)
                item.TypeAttributes = TypeAttributes.Public;
            else
                item.TypeAttributes = TypeAttributes.NestedFamANDAssem;

            if (isSealed)
            {
                item.TypeAttributes |= TypeAttributes.Sealed;
            }

            item.IsPartial = isPartial;

            item.TypeAttributes |= TypeAttributes.BeforeFieldInit | TypeAttributes.Class;

            return item;
        }

        /// <summary>
        /// Creates new constant field with given name and value.
        /// </summary>
        private static CodeMemberField CreateConstField(Type type, string name, string value, bool fieldRef)
        {
            var item = new CodeMemberField(type, name)
            {
                Attributes = MemberAttributes.Const | MemberAttributes.Public
            };

            if (fieldRef)
            {
                item.InitExpression = new CodePrimitiveExpression(value);
            }
            else
            {
                item.InitExpression = new CodeSnippetExpression(value);
            }

            return item;
        }

        /// <summary>
        /// Creates new static/read-only field with given name and value.
        /// </summary>
        private static CodeMemberField CreateStaticField(Type type, string name, object value, bool fieldRef)
        {
            var item = new CodeMemberField(type, name);

            var param = fieldRef
                            ? (CodeExpression)new CodeSnippetExpression((string)value)
                            : new CodePrimitiveExpression(value);

            item.Attributes = MemberAttributes.Static | MemberAttributes.Public;
            item.InitExpression = new CodeObjectCreateExpression(type, param);

            return item;
        }

        #endregion

        #region Code Generation

        /// <summary>
        /// Generates source code from given namespace.
        /// </summary>
        private static string GenerateFromNamespace(CodeDomProvider codeProvider, System.CodeDom.CodeNamespace codeNamespace, bool blankLinesBetweenMembers, bool isCSharp)
        {
            var result = new StringBuilder();

            using (var writer = new StringWriter(result))
            {
                var options = new CodeGeneratorOptions
                {
                    BlankLinesBetweenMembers = blankLinesBetweenMembers,
                    ElseOnClosing = true,
                    VerbatimOrder = true,
                    BracingStyle = "C",
                };

                // generate the code:
                codeProvider.GenerateCodeFromNamespace(codeNamespace, writer, options);

                // send it to the StringBuilder object:
                writer.Flush();
            }

            if (isCSharp)
            {
                result.Replace(" class ", " static class ");
                result.Replace(" partial static ", " static partial ");

                result.Replace(" static System.Guid ", " static readonly System.Guid ");
            }

            return result.ToString();
        }

        /// <summary>
        /// Converts given number into hex string.
        /// </summary>
        private static string ToHex(string number, string language)
        {
            if (!string.IsNullOrEmpty(number))
            {
                uint value;

                if (uint.TryParse(number, out value))
                    return ToHex(value, language);

                if (uint.TryParse(number, NumberStyles.HexNumber, CultureInfo.CurrentCulture, out value))
                    return ToHex(value, language);

                if ((number.StartsWith("0x") || number.StartsWith("0X") || number.StartsWith("&H")) &&
                    uint.TryParse(number.Substring(2), NumberStyles.HexNumber, CultureInfo.CurrentCulture, out value))
                    return ToHex(value, language);
            }

            // parsing failed, return string:
            return number;
        }

        /// <summary>
        /// Serialize given number into hex representation.
        /// </summary>
        private static string ToHex(uint value, string language)
        {
            switch (language)
            {
                case CodeModelLanguageConstants.vsCMLanguageVB:
                    return "&H" + value.ToString("X4");

                default:
                    return "0x" + value.ToString("X4");
            }
        }

        #endregion
    }
}
