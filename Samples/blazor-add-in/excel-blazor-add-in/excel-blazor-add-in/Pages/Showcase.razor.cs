/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;

using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Versioning;

namespace BlazorAddIn.Pages
{
    [SupportedOSPlatform("browser")]
    public partial class Showcase
    {
        [Inject, AllowNull]
        private IJSRuntime JSRuntime { get; set; }
        private IJSObjectReference? JSModule { get; set; }

        private FormModel Model { get; set; } = FormModel.Default();
        private string? StatusMessage { get; set; }
        private string StatusKind { get; set; } = "info";
        private int FontSize { get; set; } = 12;

        public record HeroTheme(string Name, string Gradient);

        private static readonly HeroTheme[] HeroThemes = new[]
        {
            new HeroTheme("Ocean",  "linear-gradient(135deg, #0078d4 0%, #5b2e91 100%)"),
            new HeroTheme("Sunset", "linear-gradient(135deg, #ff6a00 0%, #ee0979 100%)"),
            new HeroTheme("Forest", "linear-gradient(135deg, #11998e 0%, #38ef7d 100%)"),
            new HeroTheme("Slate",  "linear-gradient(135deg, #232526 0%, #414345 100%)"),
        };

        private HeroTheme SelectedTheme { get; set; } = HeroThemes[0];

        private bool IsDarkMode { get; set; }

        private int RowNumber { get; set; } = 1;
        private bool RowLocked { get; set; }
        private string ColumnLetter { get; set; } = "A";
        private bool ColumnLocked { get; set; }
        private int CellRow { get; set; } = 1;
        private string CellColumn { get; set; } = "A";
        private bool CellLocked { get; set; }

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>(
                    "import", "./Pages/Showcase.razor.js");
            }
        }

        private async Task WriteToExcel()
        {
            if (JSModule is null) return;

            try
            {
                await JSModule.InvokeVoidAsync("writeFormToSheet", Model);
                StatusKind = "success";
                StatusMessage = $"Wrote {Model.Name}'s entry to the 'Showcase' worksheet.";
            }
            catch (Exception ex)
            {
                StatusKind = "error";
                StatusMessage = $"Something went wrong: {ex.Message}";
            }
        }

        private void OnFontSizeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out var value))
            {
                FontSize = value;
            }
        }

        private void OnRatingChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out var value))
            {
                Model.Rating = value;
            }
        }

        private void OnColorChanged(ChangeEventArgs e)
        {
            var value = e.Value?.ToString();
            if (!string.IsNullOrEmpty(value))
            {
                Model.Color = value;
            }
        }

        private async Task OnRowLockChanged()
        {
            if (JSModule is null || RowNumber < 1) return;
            try
            {
                await JSModule.InvokeVoidAsync("setRowLock", RowNumber, RowLocked);
                StatusKind = "success";
                StatusMessage = $"Row {RowNumber} {(RowLocked ? "locked" : "unlocked")}.";
            }
            catch (Exception ex)
            {
                StatusKind = "error";
                StatusMessage = $"Row lock failed: {ex.Message}";
            }
        }

        private async Task OnColumnLockChanged()
        {
            if (JSModule is null || string.IsNullOrWhiteSpace(ColumnLetter)) return;
            try
            {
                await JSModule.InvokeVoidAsync("setColumnLock", ColumnLetter, ColumnLocked);
                StatusKind = "success";
                StatusMessage = $"Column {ColumnLetter.ToUpper()} {(ColumnLocked ? "locked" : "unlocked")}.";
            }
            catch (Exception ex)
            {
                StatusKind = "error";
                StatusMessage = $"Column lock failed: {ex.Message}";
            }
        }

        private async Task OnCellLockChanged()
        {
            if (JSModule is null || CellRow < 1 || string.IsNullOrWhiteSpace(CellColumn)) return;
            try
            {
                await JSModule.InvokeVoidAsync("setCellLock", CellColumn, CellRow, CellLocked);
                StatusKind = "success";
                StatusMessage = $"Cell {CellColumn.ToUpper()}{CellRow} {(CellLocked ? "locked" : "unlocked")}.";
            }
            catch (Exception ex)
            {
                StatusKind = "error";
                StatusMessage = $"Cell lock failed: {ex.Message}";
            }
        }

        private void ResetForm()
        {
            Model = FormModel.Default();
            StatusMessage = null;
        }

        public class FormModel
        {
            [Required, StringLength(60, MinimumLength = 2)]
            public string Name { get; set; } = string.Empty;

            [Required, EmailAddress]
            public string Email { get; set; } = string.Empty;

            public string Notes { get; set; } = string.Empty;

            [Range(0, 120)]
            public int Age { get; set; } = 30;

            public DateTime StartDate { get; set; } = DateTime.Today;

            public string Role { get; set; } = string.Empty;

            public int Rating { get; set; } = 7;

            public string Color { get; set; } = "#0078d4";

            public string Contact { get; set; } = "Email";

            public bool Subscribe { get; set; } = true;

            public static FormModel Default() => new();
        }
    }
}
