using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarTagNameMatcher
    {
        private readonly string _structureRebarTagName;
        private readonly string _bendingDetailName;

        public RebarTagNameMatcher(string structureRebarTagName, string bendingDetailName)
        {
            _structureRebarTagName = structureRebarTagName ?? string.Empty;
            _bendingDetailName = bendingDetailName ?? string.Empty;
        }

        public bool IsStructureRebarTag(Document doc, IndependentTag tag)
        {
            return IsMatchAny(doc, tag, _structureRebarTagName);
        }

        public bool IsBendingDetailTag(Document doc, IndependentTag tag)
        {
            return IsMatchAny(doc, tag, _bendingDetailName);
        }

        private bool IsMatchAny(Document doc, IndependentTag tag, string expectedName)
        {
            if (string.IsNullOrWhiteSpace(expectedName) || tag == null)
            {
                return false;
            }

            var names = GetCandidateNames(doc, tag);
            return names.Any(n => IsNameHit(n, expectedName));
        }

        private bool IsNameHit(string candidate, string expectedName)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(expectedName))
            {
                return false;
            }

            if (string.Equals(candidate, expectedName, StringComparison.Ordinal))
            {
                return true;
            }

            return candidate.IndexOf(expectedName, StringComparison.Ordinal) >= 0;
        }

        private IReadOnlyList<string> GetCandidateNames(Document doc, IndependentTag tag)
        {
            var result = new List<string>();

            AddIfNotEmpty(result, GetElementName(tag));

            var typeName = GetTagTypeName(doc, tag);
            AddIfNotEmpty(result, typeName);

            var familyName = GetTagFamilyName(doc, tag);
            AddIfNotEmpty(result, familyName);

            return result;
        }

        private string GetElementName(Element e)
        {
            try
            {
                return e?.Name ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private string GetTagTypeName(Document doc, IndependentTag tag)
        {
            try
            {
                var type = doc?.GetElement(tag.GetTypeId()) as ElementType;
                return type?.Name ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private string GetTagFamilyName(Document doc, IndependentTag tag)
        {
            try
            {
                var type = doc?.GetElement(tag.GetTypeId()) as ElementType;
                return type?.FamilyName ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private void AddIfNotEmpty(ICollection<string> list, string value)
        {
            if (list == null || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            list.Add(value);
        }
    }
}

