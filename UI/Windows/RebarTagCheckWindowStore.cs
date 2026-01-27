using System;
using System.Diagnostics;

namespace ShimizRevitAddin2026.UI.Windows
{
    internal class RebarTagCheckWindowStore
    {
        private RebarTagCheckWindow _current;

        public RebarTagCheckWindow Current => _current;

        public bool TryActivateExisting()
        {
            if (_current == null)
            {
                return false;
            }

            try
            {
                if (_current.IsVisible)
                {
                    _current.Activate();
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        public void SetCurrent(RebarTagCheckWindow window)
        {
            _current = window;
            if (_current == null)
            {
                return;
            }

            try
            {
                _current.Closed += OnWindowClosed;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
            try
            {
                if (ReferenceEquals(sender, _current))
                {
                    _current = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

