using System.Drawing;

namespace MailPrioritizer.TaskPane
{
    /// <summary>
    /// Task Pane 색상 팔레트. Outlook 테마에 맞는 3가지 기본 테마 제공.
    /// </summary>
    internal class ThemePalette
    {
        // ── 배경 ─────────────────────────────────────────────────────────
        public Color PanelBackground   { get; set; }
        public Color ControlBackground { get; set; }

        // ── 텍스트 ───────────────────────────────────────────────────────
        public Color PrimaryText   { get; set; }
        public Color SecondaryText { get; set; }
        public Color HintText      { get; set; }

        // ── 구조 ─────────────────────────────────────────────────────────
        public Color Divider { get; set; }
        public Color Border  { get; set; }

        // ── 버튼 ─────────────────────────────────────────────────────────
        public Color ButtonBackground { get; set; }
        public Color ButtonText       { get; set; }
        public bool  ButtonFlat       { get; set; }

        // ── 우선순위 배지 배경 ────────────────────────────────────────────
        public Color UrgentBadge { get; set; }
        public Color HighBadge   { get; set; }
        public Color NormalBadge { get; set; }
        public Color LowBadge    { get; set; }

        // ── 우선순위 배지 텍스트 ──────────────────────────────────────────
        public Color UrgentBadgeText { get; set; }
        public Color HighBadgeText   { get; set; }
        public Color NormalBadgeText { get; set; }
        public Color LowBadgeText    { get; set; }

        // ── 상태 색 ──────────────────────────────────────────────────────
        public Color AnalyzingText { get; set; }
        public Color ComboBack     { get; set; }
        public Color ComboText     { get; set; }

        // ── 기본 테마 (Outlook 밝은/White 모드) ───────────────────────────
        public static readonly ThemePalette Light = new ThemePalette
        {
            PanelBackground   = Color.White,
            ControlBackground = Color.FromArgb(250, 250, 250),
            PrimaryText       = Color.FromArgb(26, 26, 26),
            SecondaryText     = Color.FromArgb(80, 80, 80),
            HintText          = Color.FromArgb(140, 140, 140),
            Divider           = Color.FromArgb(208, 208, 208),
            Border            = Color.FromArgb(180, 180, 180),
            ButtonBackground  = SystemColors.Control,
            ButtonText        = SystemColors.ControlText,
            ButtonFlat        = false,
            ComboBack         = Color.White,
            ComboText         = Color.FromArgb(26, 26, 26),
            UrgentBadge       = Color.FromArgb(255, 232, 232),
            HighBadge         = Color.FromArgb(255, 243, 224),
            NormalBadge       = Color.FromArgb(232, 245, 233),
            LowBadge          = Color.FromArgb(245, 245, 245),
            UrgentBadgeText   = Color.FromArgb(160, 0,   0),
            HighBadgeText     = Color.FromArgb(130, 70,  0),
            NormalBadgeText   = Color.FromArgb(0,   100, 0),
            LowBadgeText      = Color.FromArgb(90,  90,  90),
            AnalyzingText     = Color.CornflowerBlue,
        };

        // ── 회색 테마 (Outlook Dark Grey / 회색 모드) ─────────────────────
        public static readonly ThemePalette Grey = new ThemePalette
        {
            PanelBackground   = Color.FromArgb(224, 224, 224),
            ControlBackground = Color.FromArgb(207, 207, 207),
            PrimaryText       = Color.FromArgb(20, 20, 20),
            SecondaryText     = Color.FromArgb(55, 55, 55),
            HintText          = Color.FromArgb(95, 95, 95),
            Divider           = Color.FromArgb(155, 155, 155),
            Border            = Color.FromArgb(135, 135, 135),
            ButtonBackground  = Color.FromArgb(195, 195, 195),
            ButtonText        = Color.FromArgb(20, 20, 20),
            ButtonFlat        = true,
            ComboBack         = Color.FromArgb(207, 207, 207),
            ComboText         = Color.FromArgb(20, 20, 20),
            UrgentBadge       = Color.FromArgb(196, 50,  50),
            HighBadge         = Color.FromArgb(185, 110,  0),
            NormalBadge       = Color.FromArgb(46,  125, 50),
            LowBadge          = Color.FromArgb(115, 115, 115),
            UrgentBadgeText   = Color.White,
            HighBadgeText     = Color.White,
            NormalBadgeText   = Color.White,
            LowBadgeText      = Color.White,
            AnalyzingText     = Color.FromArgb(60, 120, 200),
        };

        // ── 어두운 테마 (Outlook Black 모드) ──────────────────────────────
        public static readonly ThemePalette Dark = new ThemePalette
        {
            PanelBackground   = Color.FromArgb(37, 37, 37),
            ControlBackground = Color.FromArgb(48, 48, 48),
            PrimaryText       = Color.FromArgb(220, 220, 220),
            SecondaryText     = Color.FromArgb(170, 170, 170),
            HintText          = Color.FromArgb(120, 120, 120),
            Divider           = Color.FromArgb(70,  70,  70),
            Border            = Color.FromArgb(85,  85,  85),
            ButtonBackground  = Color.FromArgb(58,  58,  58),
            ButtonText        = Color.FromArgb(215, 215, 215),
            ButtonFlat        = true,
            ComboBack         = Color.FromArgb(48,  48,  48),
            ComboText         = Color.FromArgb(220, 220, 220),
            UrgentBadge       = Color.FromArgb(160, 30,  30),
            HighBadge         = Color.FromArgb(165, 80,   0),
            NormalBadge       = Color.FromArgb(35,  100, 38),
            LowBadge          = Color.FromArgb(62,  62,  62),
            UrgentBadgeText   = Color.FromArgb(255, 200, 200),
            HighBadgeText     = Color.FromArgb(255, 220, 160),
            NormalBadgeText   = Color.FromArgb(180, 255, 180),
            LowBadgeText      = Color.FromArgb(180, 180, 180),
            AnalyzingText     = Color.FromArgb(120, 170, 255),
        };

        // ── 팩토리 ───────────────────────────────────────────────────────
        public static ThemePalette FromName(string name)
        {
            switch ((name ?? string.Empty).ToLowerInvariant())
            {
                case "grey": return Grey;
                case "dark": return Dark;
                default:     return Light;
            }
        }

        public string DisplayName
        {
            get
            {
                if (this == Grey) return "회색 테마";
                if (this == Dark) return "어두운 테마";
                return "밝은 테마";
            }
        }
    }
}
