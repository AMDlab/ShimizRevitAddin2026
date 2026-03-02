using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.ExternalEvents
{
    internal class RebarTagHighlightExternalEventHandler : IExternalEventHandler
    {
        private readonly RebarTagHighlighter _highlighter;
        private readonly RebarTagLeaderBendingDetailConsistencyService _consistencyService;

        private ElementId _rebarId = ElementId.InvalidElementId;
        private ElementId _viewId = ElementId.InvalidElementId;
        private Action<RebarTagCheckResult> _onCompleted;

        private ElementId _leaderEndMarkerCurveId = ElementId.InvalidElementId;

        public RebarTagHighlightExternalEventHandler(
            RebarTagHighlighter highlighter,
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            _highlighter = highlighter;
            _consistencyService = consistencyService;
        }

        public void SetRequest(ElementId rebarId, ElementId viewId, Action<RebarTagCheckResult> onCompleted)
        {
            _rebarId = rebarId ?? ElementId.InvalidElementId;
            _viewId = viewId ?? ElementId.InvalidElementId;
            _onCompleted = onCompleted;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app?.ActiveUIDocument;
                var doc = uidoc?.Document;
                if (uidoc == null || doc == null)
                {
                    NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>(), string.Empty));
                    return;
                }

                var (rebar, view) = ResolveTargets(doc);
                if (rebar == null || view == null)
                {
                    NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>(), string.Empty));
                    return;
                }

                var items = GetConsistencyItemsOrEmpty(doc, rebar, view);
                var extraIds = BuildExtraHighlightIds(items);

                var model = _highlighter.Highlight(uidoc, rebar, view, extraIds);
                var message = FormatMessage(items);
                var result = BuildResult(doc, model, message);

                TryRedrawLeaderEndMarker(doc, view, rebar);
                NotifyCompleted(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                try
                {
                    TaskDialog.Show("RebarTag", ex.ToString());
                }
                catch (Exception dialogEx)
                {
                    Debug.WriteLine(dialogEx);
                }

                NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>(), string.Empty));
            }
        }

        public string GetName()
        {
            return "RebarTagHighlightExternalEventHandler";
        }

        private (Rebar rebar, View view) ResolveTargets(Document doc)
        {
            var rebar = doc.GetElement(_rebarId) as Rebar;
            var view = doc.GetElement(_viewId) as View;
            return (rebar, view);
        }

        private RebarTagCheckResult BuildResult(Document doc, ShimizRevitAddin2026.Model.RebarTag model, string message)
        {
            if (doc == null || model == null)
            {
                return new RebarTagCheckResult(new List<string>(), new List<string>(), string.Empty);
            }

            var structure = BuildTagContentList(doc, model.StructureTagIds);
            var bending = BuildTagContentList(doc, model.BendingDetailTagIds);
            return new RebarTagCheckResult(structure, bending, message);
        }

        private IReadOnlyList<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem> GetConsistencyItemsOrEmpty(
            Document doc,
            Rebar rebar,
            View view)
        {
            try
            {
                if (_consistencyService == null || doc == null || rebar == null || view == null)
                {
                    return new List<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem>();
                }

                return _consistencyService.Check(doc, rebar, view)
                    ?? new List<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem>();
            }
        }

        private string FormatMessage(IReadOnlyList<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return "曲げ詳細ID: （未取得）\n判定結果を取得できません。";
            }

            // 現状は1件（自由端タグの判定結果）を返す
            var item = items.FirstOrDefault(x => x != null);
            if (item == null)
            {
                return "曲げ詳細ID: （未取得）\n判定結果を取得できません。";
            }

            var lines = new List<string>();
            lines.Add(BuildPointedBendingDetailLine(item));

            // OK/NG に関係なく詳細メッセージを表示する（空なら省略）
            var msg = item.Message ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(msg))
            {
                lines.Add(msg);
            }

            return string.Join(Environment.NewLine, lines);
        }

        private IReadOnlyList<ElementId> BuildExtraHighlightIds(
            IReadOnlyList<ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem> items)
        {
            try
            {
                var result = new List<ElementId>();
                if (items == null || items.Count == 0) return result;

                var item = items.FirstOrDefault(x => x != null);
                if (item == null) return result;

                AddIfValid(result, item.PointedDirectRebarId);
                AddIfValid(result, item.PointedBendingDetailId);
                AddIfValid(result, item.PointedRebarId);

                return result.Distinct().ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private void AddIfValid(ICollection<ElementId> ids, ElementId id)
        {
            if (ids == null) return;
            if (id == null || id == ElementId.InvalidElementId) return;
            ids.Add(id);
        }

        private void TryRedrawLeaderEndMarker(Document doc, View view, Rebar rebar)
        {
            try
            {
                if (_consistencyService == null) return;
                if (doc == null || view == null || rebar == null) return;

                var (ok, endPoint, reason) = _consistencyService.TryGetLeaderEndPoint(doc, rebar, view);
                if (!ok || endPoint == null)
                {
                    // マーカーが作れない場合でも、前回のマーカーは消す
                    DeleteLeaderEndMarker(doc);
                    if (!string.IsNullOrWhiteSpace(reason)) Debug.WriteLine(reason);
                    return;
                }

                using (var tx = new Transaction(doc, "RebarTag: Leader end marker"))
                {
                    tx.Start();
                    DeleteLeaderEndMarkerInternal(doc);
                    CreateLeaderEndMarkerInternal(doc, view, endPoint);
                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void DeleteLeaderEndMarker(Document doc)
        {
            try
            {
                if (doc == null) return;
                using (var tx = new Transaction(doc, "RebarTag: Delete leader end marker"))
                {
                    tx.Start();
                    DeleteLeaderEndMarkerInternal(doc);
                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _leaderEndMarkerCurveId = ElementId.InvalidElementId;
            }
        }

        private void DeleteLeaderEndMarkerInternal(Document doc)
        {
            if (doc == null) return;
            if (_leaderEndMarkerCurveId == null || _leaderEndMarkerCurveId == ElementId.InvalidElementId) return;

            try
            {
                if (doc.GetElement(_leaderEndMarkerCurveId) != null)
                    doc.Delete(_leaderEndMarkerCurveId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                _leaderEndMarkerCurveId = ElementId.InvalidElementId;
            }
        }

        private void CreateLeaderEndMarkerInternal(Document doc, View view, XYZ center)
        {
            try
            {
                if (doc == null || view == null || center == null) return;

                var arc = BuildLeaderEndCircle(view, center, 150.0 / 304.8); // 150mm
                if (arc == null) return;

                var curve = doc.Create.NewDetailCurve(view, arc);
                if (curve == null) return;

                _leaderEndMarkerCurveId = curve.Id;
                ApplyLeaderEndMarkerGraphics(view, _leaderEndMarkerCurveId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _leaderEndMarkerCurveId = ElementId.InvalidElementId;
            }
        }

        private Arc BuildLeaderEndCircle(View view, XYZ center, double radiusFeet)
        {
            try
            {
                if (view == null || center == null) return null;
                if (radiusFeet <= 0) return null;

                var x = view.RightDirection;
                var y = view.UpDirection;
                if (x == null || y == null) return null;

                return Arc.Create(center, radiusFeet, 0, 2 * Math.PI, x, y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
        }

        private void ApplyLeaderEndMarkerGraphics(View view, ElementId id)
        {
            try
            {
                if (view == null) return;
                if (id == null || id == ElementId.InvalidElementId) return;

                var ogs = new OverrideGraphicSettings();
                ogs.SetProjectionLineColor(new Color(0, 200, 255));
                view.SetElementOverrides(id, ogs);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private string BuildPointedBendingDetailLine(ShimizRevitAddin2026.Model.RebarTagLeaderBendingDetailCheckItem item)
        {
            if (item == null)
            {
                return "曲げ詳細ID: （未取得）";
            }

            if (item.PointedBendingDetailId == null || item.PointedBendingDetailId == ElementId.InvalidElementId)
            {
                return "曲げ詳細ID: （未取得）";
            }

            return $"曲げ詳細ID: {item.PointedBendingDetailId.Value}";
        }

        private IReadOnlyList<string> BuildTagContentList(Document doc, IReadOnlyList<ElementId> ids)
        {
            if (doc == null || ids == null)
            {
                return new List<string>();
            }

            return ids
                .Select(id => BuildTagContent(doc, id))
                .Where(x => x != null)
                .ToList();
        }

        private string BuildTagContent(Document doc, ElementId id)
        {
            if (doc == null || id == null)
            {
                return string.Empty;
            }

            try
            {
                var e = doc.GetElement(id);
                if (e is IndependentTag tag)
                {
                    return BuildTagTextWithId(tag.TagText, id);
                }

                return id.Value.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return id.Value.ToString();
            }
        }

        private string BuildTagTextWithId(string content, ElementId id)
        {
            if (id == null)
            {
                return content ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                return id.Value.ToString();
            }

            // タグ文字と要素IDを同時に表示する
            return $"{content} / {id.Value}";
        }

        private void NotifyCompleted(RebarTagCheckResult result)
        {
            try
            {
                _onCompleted?.Invoke(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

