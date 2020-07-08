// --------------------------------------------------------------------------------------------------------------------
// <copyright company="twentysix" file="BoolExtensions.cs">
// Copyright (c) twentysix.  All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace ExcelPluginTest.Extensions
{
    public static class BoolExtensions
    {
        public static string ToYesNo(this bool value)
        {
            return value ? "Yes" : "No";
        }
    }
}