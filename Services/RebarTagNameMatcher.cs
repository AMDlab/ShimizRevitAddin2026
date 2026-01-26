using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarTagNameMatcher
    {
        private readonly string _bendingDetailName;

        public RebarTagNameMatcher(string bendingDetailName)
        {
            _bendingDetailName = bendingDetailName ?? string.Empty;
        }

        // アクティブビュー内の「構造鉄筋タグ」判定
        public bool IsStructureRebarTag(Document doc, Element tagElement)
        {
            return IsBuiltInCategory(tagElement, BuiltInCategory.OST_RebarTags);
        }

        // アクティブビュー内の「曲げ加工詳細」判定
        public bool IsBendingDetailTag(Document doc, Element tagElement)
        {
            return IsMatchAny(doc, tagElement, _bendingDetailName);
        }

        private bool IsMatchAny(Document doc, Element tagElement, string expectedName)
        {
            if (string.IsNullOrWhiteSpace(expectedName) || tagElement == null)
            {
                return false;
            }

            var names = GetCandidateNames(doc, tagElement);
            return names.Any(n => IsNameHit(n, expectedName));
        }

        private bool IsNameHit(string candidate, string expectedName)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(expectedName))
            {
                return false;
            }

            var c = Normalize(candidate);
            var e = Normalize(expectedName);
            if (string.IsNullOrWhiteSpace(c) || string.IsNullOrWhiteSpace(e))
            {
                return false;
            }

            if (string.Equals(c, e, StringComparison.Ordinal))
            {
                return true;
            }

            return c.IndexOf(e, StringComparison.Ordinal) >= 0;
        }

        private IReadOnlyList<string> GetCandidateNames(Document doc, Element tagElement)
        {
            var result = new List<string>();

            AddIfNotEmpty(result, GetElementName(tagElement));

            var typeName = GetTypeName(doc, tagElement);
            AddIfNotEmpty(result, typeName);

            var familyName = GetFamilyName(doc, tagElement);
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

        private string GetTypeName(Document doc, Element tagElement)
        {
            try
            {
                var type = doc?.GetElement(tagElement.GetTypeId()) as ElementType;
                return type?.Name ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private string GetFamilyName(Document doc, Element tagElement)
        {
            try
            {
                var type = doc?.GetElement(tagElement.GetTypeId()) as ElementType;
                return type?.FamilyName ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private bool IsBuiltInCategory(Element e, BuiltInCategory bic)
        {
            try
            {
                if (e?.Category?.Id == null)
                {
                    return false;
                }

                // Revit 2026: ElementId.Value を利用
                return e.Category.Id.Value == (long)(int)bic;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        private string Normalize(string s)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    return string.Empty;
                }

                // 全角/半角・互換文字を吸収して、空白も除去する
                var normalized = s.Normalize(NormalizationForm.FormKC).Trim();
                var sb = new StringBuilder(normalized.Length);
                foreach (var ch in normalized)
                {
                    if (char.IsWhiteSpace(ch))
                    {
                        continue;
                    }

                    sb.Append(ch);
                }

                return sb.ToString();
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

