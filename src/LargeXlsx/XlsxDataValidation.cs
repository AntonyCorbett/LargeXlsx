using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace LargeXlsx
{
    public class XlsxDataValidation : IEquatable<XlsxDataValidation>
    {
        public bool AllowBlank { get; }
        public string Error { get; }
        public string ErrorTitle { get; }
        public ErrorStyle? ErrorStyleValue { get; }
        public Operator? OperatorValue { get; }
        public string Prompt { get; }
        public string PromptTitle { get; }
        public bool ShowDropDown { get; }
        public bool ShowErrorMessage { get; }
        public bool ShowInputMessage { get; }
        public ValidationType? ValidationTypeValue { get; }
        public string Formula1 { get; }
        public string Formula2 { get; }

        public enum ErrorStyle { Information, Stop, Warning }
        public enum Operator { Between, Equal, GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, NotBetween, NotEqual }
        public enum ValidationType { Custom, Date, Decimal, List, None, TextLength, Time, Whole }

        public XlsxDataValidation(
            bool allowBlank = false,
            string error = null,
            string errorTitle = null,
            ErrorStyle? errorStyle = null,
            Operator? operatorType = null,
            string prompt = null,
            string promptTitle = null,
            bool showDropDown = false,
            bool showErrorMessage = false,
            bool showInputMessage = false,
            ValidationType? validationType = null,
            string formula1 = null,
            string formula2 = null)
        {
            AllowBlank = allowBlank;
            Error = error;
            ErrorTitle = errorTitle;
            ErrorStyleValue = errorStyle;
            OperatorValue = operatorType;
            Prompt = prompt;
            PromptTitle = promptTitle;
            ShowDropDown = showDropDown;
            ShowErrorMessage = showErrorMessage;
            ShowInputMessage = showInputMessage;
            ValidationTypeValue = validationType;
            Formula1 = formula1;
            Formula2 = formula2;
        }

        public static XlsxDataValidation List(
            IEnumerable<string> choices,
            bool allowBlank = false,
            string error = null,
            string errorTitle = null,
            ErrorStyle? errorStyle = null,
            string prompt = null,
            string promptTitle = null,
            bool showDropDown = false,
            bool showErrorMessage = false,
            bool showInputMessage = false)
        {
            return new XlsxDataValidation(allowBlank, error, errorTitle, errorStyle, null, prompt, promptTitle, showDropDown, showErrorMessage, showInputMessage,
                ValidationType.List, '"'+ string.Join(",", choices.Select(c => c.Replace("\"", "\"\""))) + '"');
        }

        #region Equality members
        public bool Equals(XlsxDataValidation other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return AllowBlank == other.AllowBlank && Error == other.Error && ErrorTitle == other.ErrorTitle
                   && ErrorStyleValue == other.ErrorStyleValue && OperatorValue == other.OperatorValue
                   && Prompt == other.Prompt && PromptTitle == other.PromptTitle && ShowDropDown == other.ShowDropDown
                   && ShowErrorMessage == other.ShowErrorMessage && ShowInputMessage == other.ShowInputMessage
                   && ValidationTypeValue == other.ValidationTypeValue
                   && Formula1 == other.Formula1 && Formula2 == other.Formula2;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((XlsxDataValidation)obj);
        }

        public override int GetHashCode()
        {
            return (
                AllowBlank,
                Error,
                ErrorTitle,
                ErrorStyleValue,
                OperatorValue,
                Prompt,
                PromptTitle,
                ShowDropDown,
                ShowErrorMessage,
                ShowInputMessage,
                ValidationTypeValue,
                Formula1,
                Formula2).GetHashCode();
        }

        public static bool operator ==(XlsxDataValidation left, XlsxDataValidation right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxDataValidation left, XlsxDataValidation right)
        {
            return !Equals(left, right);
        }
        #endregion
    }
}