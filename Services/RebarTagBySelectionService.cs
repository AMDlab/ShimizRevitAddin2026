using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace ShimizRevitAddin2026.Services
{
    /// <summary>
    /// 選択した鉄筋（個別バー）に「Tag rebar_2」族・タイプ「標準」でタグを付与するサービス。
    /// 重要：Rebar Set 全体ではなく、必ず individual bar（subelement）を選択させる。
    /// </summary>
    internal class RebarTagBySelectionService
    {
        private const string FamilyNamePattern = "Tag rebar_2";
        private const string TypeNamePattern = "標準";

        /// <summary>
        /// 指定ビュー内で、選択された鉄筋（Reference と Rebar のペア）にタグを付与する。
        /// </summary>
        public (int taggedCount, string errorMessage) TagRebarsInView(
            Document doc,
            View view,
            IReadOnlyList<(Reference Reference, Rebar Rebar)> selection)
        {
            if (doc == null) return (0, "Document が null です。");
            if (view == null) return (0, "View が null です。");
            if (selection == null || selection.Count == 0) return (0, "鉄筋が選択されていません。");

            var (hasSymbol, symbolId, symbolError) = TryFindRebarTagSymbol(doc);
            if (!hasSymbol || symbolId == null || symbolId == ElementId.InvalidElementId)
                return (0, symbolError ?? "タグ族「Tag rebar_2」タイプ「標準」が見つかりません。");

            if (!TryActivateTagSymbol(doc, symbolId))
                return (0, "タグタイプの有効化に失敗しました。");

            var taggedCount = 0;
            var failures = new List<string>();

            foreach (var (rawRef, rebar) in selection)
            {
                if (rawRef == null || rebar == null) continue;

                // ① Reference が Rebar Set / Element 参照っぽい場合は事前に弾く
                //    （このケースは API で救済が難しいため、操作を正してもらうのが最短）
                if (!LooksLikeTaggableRebarReference(rawRef))
                {
                    failures.Add($"[{rebar.Id.Value}] Rebar Set を選択している可能性があります。TAB で individual bar を選択してください。");
                    continue;
                }

                // ② 2023+ の “選べているのに tag できない” 対策：stable rep の round-trip で reference を正規化
                var tagRef = NormalizeReferenceByStableRoundTrip(doc, rawRef) ?? rawRef;

                var (placed, msg) = TryPlaceTagForRebar(doc, view, tagRef, rebar, symbolId);
                if (placed) taggedCount++;
                else if (!string.IsNullOrWhiteSpace(msg)) failures.Add($"[{rebar.Id.Value}] {msg}");
            }

            var errorMessage = failures.Count > 0 ? string.Join(Environment.NewLine, failures.Take(8)) : null;
            if (failures.Count > 8) errorMessage += Environment.NewLine + $"他 {failures.Count - 8} 件の失敗があります。";

            return (taggedCount, errorMessage);
        }

        /// <summary>
        /// Rebar を tag できそうな reference かざっくり判定する。
        /// </summary>
        private bool LooksLikeTaggableRebarReference(Reference reference)
        {
            try
            {
                // NOTE: Rebar Set を Element として拾った reference は NONE になりがち
                //       individual bar の subelement は LINEAR/EDGE などになることが多い
                return reference.ElementReferenceType != ElementReferenceType.REFERENCE_TYPE_NONE;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// StableRepresentation を経由して Reference を再生成する（2023+ で効くことがある）。
        /// </summary>
        private Reference NormalizeReferenceByStableRoundTrip(Document doc, Reference reference)
        {
            if (doc == null || reference == null) return null;

            try
            {
                var stable = reference.ConvertToStableRepresentation(doc);
                if (string.IsNullOrWhiteSpace(stable)) return null;

                // NOTE: 文字列を Parse し直すことで内部的に taggable reference になるケースがある
                var parsed = Reference.ParseFromStableRepresentation(doc, stable);
                return parsed;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
        }

        /// <summary>
        /// タグ族「Tag rebar_2」タイプ「標準」の FamilySymbol の ElementId を取得する。
        /// </summary>
        private (bool ok, ElementId symbolId, string errorMessage) TryFindRebarTagSymbol(Document doc)
        {
            try
            {
                var symbolId = FindRebarTagSymbolId(doc);
                if (symbolId != null && symbolId != ElementId.InvalidElementId)
                    return (true, symbolId, null);

                return (false, null, "タグ族「Tag rebar_2」タイプ「標準」がプロジェクトにありません。");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private ElementId FindRebarTagSymbolId(Document doc)
        {
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RebarTags)
                .WhereElementIsElementType();

            foreach (var element in collector)
            {
                var symbol = element as FamilySymbol;
                if (symbol == null) continue;

                var familyName = GetFamilyName(symbol);
                var typeName = symbol.Name ?? string.Empty;

                if (IsFamilyNameMatch(familyName, FamilyNamePattern) && IsTypeNameMatch(typeName, TypeNamePattern))
                    return symbol.Id;
            }

            return null;
        }

        private string GetFamilyName(FamilySymbol symbol)
        {
            if (symbol == null) return string.Empty;

            try
            {
                return symbol.Family?.Name ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private bool IsFamilyNameMatch(string candidate, string expected)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(expected)) return false;

            var c = candidate.Trim();
            var e = expected.Trim();

            return string.Equals(c, e, StringComparison.OrdinalIgnoreCase)
                   || c.IndexOf(e, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool IsTypeNameMatch(string candidate, string expected)
        {
            if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(expected)) return false;
            return string.Equals(candidate.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// タグの FamilySymbol を有効化する。Create 前に必要。
        /// </summary>
        private bool TryActivateTagSymbol(Document doc, ElementId symbolId)
        {
            if (doc == null || symbolId == null || symbolId == ElementId.InvalidElementId) return false;

            try
            {
                var symbol = doc.GetElement(symbolId) as FamilySymbol;
                if (symbol == null) return false;

                if (!symbol.IsActive) symbol.Activate();
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        /// <summary>
        /// 1本の鉄筋にタグを1つ付与する。
        /// </summary>
        private (bool placed, string errorMessage) TryPlaceTagForRebar(
            Document doc,
            View view,
            Reference reference,
            Rebar rebar,
            ElementId tagSymbolId)
        {
            try
            {
                var (hasPoint, leaderEndPoint) = TryGetRebarEndPoint(rebar);
                if (!hasPoint || leaderEndPoint == null)
                    return (false, "鉄筋の端部座標を取得できません。");

                // NOTE: 指定シンボル直指定で落ちるケースがあるため、まずはカテゴリで作成して後で型を切り替える
                IndependentTag tag = null;

                try
                {
                    tag = IndependentTag.Create(
                        doc,
                        view.Id,
                        reference,
                        addLeader: true,
                        TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal,
                        leaderEndPoint);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException aex)
                {
                    // NOTE: “The reference can not be tagged” がここで来る
                    Debug.WriteLine(aex);

                    return (false,
                        "The reference can not be tagged。Rebar Set を選択している可能性があります。TAB で individual bar を選択してください。");
                }

                if (tag == null)
                    return (false, "タグの作成に失敗しました。");

                // 作成後に希望の族タイプ（Tag rebar_2 - 標準）に切り替え
                if (tagSymbolId != null && tagSymbolId != ElementId.InvalidElementId && tag.GetTypeId() != tagSymbolId)
                {
                    try { tag.ChangeTypeId(tagSymbolId); }
                    catch (Exception exChange)
                    {
                        // NOTE: タイプ変更に失敗してもタグは作成済みなので成功扱い
                        Debug.WriteLine(exChange);
                    }
                }

                return (true, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// 鉄筋中心線の端部（最後の曲線の終点）を返す。
        /// </summary>
        private (bool ok, XYZ endPoint) TryGetRebarEndPoint(Rebar rebar)
        {
            if (rebar == null) return (false, null);

            try
            {
                var curves = GetRebarCenterlineCurves(rebar);
                if (curves == null || curves.Count == 0) return (false, null);

                var lastCurve = curves[curves.Count - 1];
                if (lastCurve == null) return (false, null);

                var endPoint = lastCurve.GetEndPoint(1);
                return endPoint != null ? (true, endPoint) : (false, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null);
            }
        }

        private IReadOnlyList<Curve> GetRebarCenterlineCurves(Rebar rebar)
        {
            if (rebar == null) return new List<Curve>();

            try
            {
                var curves = rebar.GetCenterlineCurves(
                    false,
                    false,
                    false,
                    MultiplanarOption.IncludeOnlyPlanarCurves,
                    0);

                return curves?.Where(c => c != null).ToList() ?? new List<Curve>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<Curve>();
            }
        }
    }
}