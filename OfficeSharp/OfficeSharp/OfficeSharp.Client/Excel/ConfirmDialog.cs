using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.LightSwitch.Runtime.Shell.Framework;

namespace OfficeSharp
{
    [TemplatePart(Name = ConfirmDialog.OkButtonElement, Type = typeof(Button))]
    [TemplatePart(Name = ConfirmDialog.CancelButtonElement, Type = typeof(Button))]
    public class ConfirmDialog : ScreenChildWindowContent
    {

        public const string OkButtonElement = "OKButton";
        public const string CancelButtonElement = "CancelButton";

        private Button withEventsField_okButton;
        private Button okButton
        {
            get { return withEventsField_okButton; }
            set
            {
                if (withEventsField_okButton != null)
                    withEventsField_okButton.Click -= okButton_Click;

                withEventsField_okButton = value;
                if (withEventsField_okButton != null)
                    withEventsField_okButton.Click += okButton_Click;
            }
        }

        private Button withEventsField_cancelButton;
        private Button cancelButton
        {
            get { return withEventsField_cancelButton; }
            set
            {
                if (withEventsField_cancelButton != null)
                    withEventsField_cancelButton.Click -= cancelButton_Click;

                withEventsField_cancelButton = value;
                if (withEventsField_cancelButton != null)
                    withEventsField_cancelButton.Click += cancelButton_Click;
            }

        }

        public ConfirmDialog()
            : base()
        {
            this.DefaultStyleKey = typeof(ConfirmDialog);
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            okButton = this.GetTemplateChild(OkButtonElement) as Button;
            cancelButton = this.GetTemplateChild(CancelButtonElement) as Button;
            UpdateTitle();
            UpdateCloseButton();
        }

        private ScreenChildWindow ParentChildWindow
        {
            get { return this.Parent as ScreenChildWindow; }
        }

        private void okButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.ParentChildWindow.DialogResult = true;
            this.ParentChildWindow.DialogResultCancel = false;
        }

        private void cancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.ParentChildWindow.DialogResult = false;
            this.ParentChildWindow.DialogResultCancel = true;
        }

        #region "Dependency Properties"
        public string Title
        {
            get { return Convert.ToString(this.GetValue(TitleProperty)); }
            set { this.SetValue(TitleProperty, value); }
        }

        public static readonly DependencyProperty TitleProperty = DependencyProperty.Register("Title", typeof(string), typeof(ConfirmDialog), new PropertyMetadata(null, TitleChanged));
        public static void TitleChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var dialog = d as ConfirmDialog;
            dialog.UpdateTitle();
        }

        private void UpdateTitle()
        {
            if (this.ParentChildWindow != null & Title != null)
            {
                this.ParentChildWindow.Title = this.Title;
            }
        }

        public bool ShowCloseButton
        {
            get { return Convert.ToBoolean(this.GetValue(ShowCloseButtonProperty)); }
            set { this.SetValue(ShowCloseButtonProperty, value); }
        }

        public static readonly DependencyProperty ShowCloseButtonProperty = DependencyProperty.Register("ShowCloseButton", typeof(bool), typeof(ConfirmDialog), new PropertyMetadata(ShowCloseButtonChanged));
        public static void ShowCloseButtonChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var dialog = d as ConfirmDialog;
            dialog.UpdateCloseButton();
        }

        private void UpdateCloseButton()
        {
            if (this.ParentChildWindow != null)
                this.ParentChildWindow.HasCloseButton = this.ShowCloseButton;
        }

        public bool OkButtonVisible
        {
            get { return Convert.ToBoolean(this.GetValue(OkButtonVisibleProperty)); }
            set { this.SetValue(OkButtonVisibleProperty, value); }
        }

        public static readonly DependencyProperty OkButtonVisibleProperty = DependencyProperty.Register("OkButtonVisible", typeof(bool), typeof(ConfirmDialog), new PropertyMetadata(true));
        public bool CancelButtonVisible
        {
            get { return Convert.ToBoolean(this.GetValue(CancelButtonVisibleProperty)); }
            set { this.SetValue(CancelButtonVisibleProperty, value); }
        }

        public static readonly DependencyProperty CancelButtonVisibleProperty = DependencyProperty.Register("CancelButtonVisible", typeof(bool), typeof(ConfirmDialog), new PropertyMetadata(true));
        public string OkButtonTitle
        {
            get { return Convert.ToBoolean(this.GetValue(OkButtonTitleProperty)).ToString(); }
            set { this.SetValue(OkButtonTitleProperty, value); }
        }

        public static readonly DependencyProperty OkButtonTitleProperty = DependencyProperty.Register("OkButtonTitle", typeof(string), typeof(ConfirmDialog), new PropertyMetadata("Ok"));
        public string CancelButtonTitle
        {
            get { return Convert.ToBoolean(this.GetValue(CancelButtonTitleProperty)).ToString(); }
            set { this.SetValue(CancelButtonTitleProperty, value); }
        }

        public static readonly DependencyProperty CancelButtonTitleProperty = DependencyProperty.Register("CancelButtonTitle", typeof(string), typeof(ConfirmDialog), new PropertyMetadata("Cancel"));
        #endregion
    }

    public class BooleanToVisibilityConverter : IValueConverter
    {

        public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool vis = (Boolean)value;
            if (vis)
                return Visibility.Visible;
            else
                return Visibility.Collapsed;
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

