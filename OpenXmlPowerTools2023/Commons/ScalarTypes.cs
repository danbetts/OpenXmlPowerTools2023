// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.ObjectModel;

namespace OpenXmlPowerTools2023.Commons
{
    internal static class DefaultScalarTypes
    {
        private static readonly Hashtable defaultScalarTypesHash;
        internal static bool IsTypeInList(Collection<string> typeNames)
        {
            string text = PSObjectIsOfExactType(typeNames);
            return !string.IsNullOrEmpty(text) && (PSObjectIsEnum(typeNames) || defaultScalarTypesHash.ContainsKey(text));
        }

        static DefaultScalarTypes()
        {
            defaultScalarTypesHash = new Hashtable(StringComparer.OrdinalIgnoreCase);
            defaultScalarTypesHash.Add("System.String", null);
            defaultScalarTypesHash.Add("System.SByte", null);
            defaultScalarTypesHash.Add("System.Byte", null);
            defaultScalarTypesHash.Add("System.Int16", null);
            defaultScalarTypesHash.Add("System.UInt16", null);
            defaultScalarTypesHash.Add("System.Int32", 10);
            defaultScalarTypesHash.Add("System.UInt32", 10);
            defaultScalarTypesHash.Add("System.Int64", null);
            defaultScalarTypesHash.Add("System.UInt64", null);
            defaultScalarTypesHash.Add("System.Char", 1);
            defaultScalarTypesHash.Add("System.Single", null);
            defaultScalarTypesHash.Add("System.Double", null);
            defaultScalarTypesHash.Add("System.Boolean", 5);
            defaultScalarTypesHash.Add("System.Decimal", null);
            defaultScalarTypesHash.Add("System.IntPtr", null);
            defaultScalarTypesHash.Add("System.Security.SecureString", null);
        }

        internal static string PSObjectIsOfExactType(Collection<string> typeNames)
        {
            if (typeNames.Count != 0)
            {
                return typeNames[0];
            }
            return null;
        }

        internal static bool PSObjectIsEnum(Collection<string> typeNames)
        {
            return typeNames.Count >= 2 && !string.IsNullOrEmpty(typeNames[1]) && string.Equals(typeNames[1], "System.Enum", StringComparison.Ordinal);
        }
    }
}
