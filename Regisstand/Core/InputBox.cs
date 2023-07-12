using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace Regisstand.Core
{
    public class InputBox : IDisposable
    {
        private Window m_Window;
        private readonly Dictionary<int, TextBox> m_TextBoxes;
        private readonly List<int> m_RequiredFieldIds;
        private bool m_ClickedOk;
        private Button m_CancelButton;
        private Button m_OkButton;
        #region [Ctor]
        public InputBox(string Title, Dictionary<int, string> RequestedValueMessages, List<int> RequiredFieldIds)
        {
            m_RequiredFieldIds = RequiredFieldIds ?? new List<int>();
            m_TextBoxes = new Dictionary<int, TextBox>();
            m_ClickedOk = false;
            OutputValues = new Dictionary<int, string>();
            __CreateWindow(Title, RequestedValueMessages);
        }
        #endregion

        #region - properties -
        #region - public properties -
        public Dictionary<int, string> OutputValues { get; set; }
        #endregion
        #endregion

        #region - public methods -
        #region [OpenInputDialog]
        public bool OpenInputDialog()
        {
            if (m_TextBoxes != null && m_TextBoxes.Any(t => t.Value != null))
            {
                var tb = m_TextBoxes.FirstOrDefault().Value;
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Input,
                  new Action(() =>
                  {
                      tb.Focus();
                      Keyboard.Focus(tb);
                  }));
            }
            m_Window.ShowDialog();
            return m_ClickedOk;
        }
        #endregion

        #region [Dispose]
        protected virtual void Dispose(bool Disposing)
        {
            if (Disposing)
            {
                m_CancelButton.Click -= __ButtonCancel_Clicked;
                m_OkButton.Click -= __ButtonOk_Clicked;

                m_TextBoxes.ToList().ForEach(tb => tb.Value.TextChanged -= __TextBox_TextChanged);
                m_Window = null;
            }
        }

        public void Dispose()
        {
            Dispose(Disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #endregion

        #region - private methods -

        #region [__CreateWindow]
        private void __CreateWindow(string Title, Dictionary<int, string> RequestedValueMessages)
        {
            m_Window = new Window()
            {
                Width = 500,
                Height = 80 + RequestedValueMessages.Count * 25,
                MinWidth = 500,
                MinHeight = 80 + RequestedValueMessages.Count * 25,
                Title = Title,
            };
            var DlgContentControl = new ContentControl();

            var MainGrid = __CreateMainGrid(RequestedValueMessages);

            DlgContentControl.Content = MainGrid;
            m_Window.Content = DlgContentControl;
        }
        #endregion

        #region [__CreateMainGrid]
        private Grid __CreateMainGrid(Dictionary<int, string> RequestedValueMessages)
        {
            var MainGrid = new Grid();
            MainGrid.Margin = new Thickness(5);
            var ButtonStackPanel = new StackPanel()
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 5, 0, 0),
            };
            var OkText = "Ok";
            var CancelText = "Abbrechen";
            m_OkButton = new Button { Height = 22, Width = 80, Content = OkText, Margin = new Thickness(0, 0, 5, 0) };
            m_OkButton.Click += __ButtonOk_Clicked;
            ButtonStackPanel.Children.Add(m_OkButton);
            m_CancelButton = new Button { Height = 22, Width = 80, Content = CancelText };
            m_CancelButton.Click += __ButtonCancel_Clicked;
            ButtonStackPanel.Children.Add(m_CancelButton);
            Grid.SetRow(ButtonStackPanel, 1);
            if (m_RequiredFieldIds.Any())
                m_OkButton.IsEnabled = false;

            var InputGrid = new Grid();
            InputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            InputGrid.ColumnDefinitions.Add(new ColumnDefinition { MaxWidth = 400 });

            int RowCounter = 0;
            foreach (var RequestedMessage in RequestedValueMessages)
            {
                InputGrid.RowDefinitions.Add(new RowDefinition());
                var DlgLabel = new Label
                {
                    Content = RequestedMessage.Value,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, -2, 5, 0)
                };
                InputGrid.Children.Add(DlgLabel);
                Grid.SetRow(DlgLabel, RowCounter);
                Grid.SetColumn(DlgLabel, 0);

                var DlgTextBox = new TextBox
                {
                    VerticalAlignment = VerticalAlignment.Top
                };
                DlgTextBox.TextChanged += __TextBox_TextChanged;
                m_TextBoxes.Add(RequestedMessage.Key, DlgTextBox);
                InputGrid.Children.Add(DlgTextBox);
                Grid.SetRow(DlgTextBox, RowCounter);
                Grid.SetColumn(DlgTextBox, 1);
                RowCounter++;
            }
            MainGrid.Children.Add(InputGrid);
            MainGrid.Children.Add(ButtonStackPanel);
            Grid.SetRow(InputGrid, 0);
            return MainGrid;
        }
        #endregion

        #region [__ButtonOk_Clicked]
        private void __ButtonOk_Clicked(object sender, EventArgs e)
        {
            m_ClickedOk = true;
            m_TextBoxes.ToList().ForEach(tb => OutputValues.Add(tb.Key, tb.Value.Text));
            m_Window.Close();
        }
        #endregion

        #region [__ButtonCancel_Clicked]
        private void __ButtonCancel_Clicked(object sender, EventArgs e)
        {
            m_ClickedOk = false;
            m_Window.Close();
        }
        #endregion

        #region [__TextBox_TextChanged]
        private void __TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var RequiredFields = m_TextBoxes.Where(x => m_RequiredFieldIds.Contains(x.Key));
            if (RequiredFields.Any())
                m_OkButton.IsEnabled = RequiredFields.All(y => !string.IsNullOrEmpty(y.Value.Text));
        }
        #endregion

        #endregion
    }
}
