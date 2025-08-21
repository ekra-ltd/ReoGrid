using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media;

namespace unvell.ReoGrid.WPF
{
    /// <summary>
    /// Класс содержащий загруженные шрифты
    /// </summary>
    public static class FontLibrary
    {
        #region Публичные поля и методы

        /// <summary>
        /// Удаление шрифта
        /// </summary>
        /// <param name="name">Имя шрифта</param>
        public static void RemoveFont(string name)
        {
            FontsDictionary.Remove(name);
            FontCollectionChanged?.Invoke(null, new FontCollectionChangedArg(name, FontCollectionChangedArg.EventType.Delete));
        }

        /// <summary>
        /// Добавление шрифта
        /// </summary>
        /// <param name="name">Имя шрифта</param>
        /// <param name="fontFamily">шрифт</param>
        public static void AddFont(string name, FontFamily fontFamily)
        {
            FontsDictionary[name] = fontFamily;
            FontCollectionChanged?.Invoke(null, new FontCollectionChangedArg(name, FontCollectionChangedArg.EventType.Add));
        }

        /// <summary>
        /// Очистка списка шрифтов
        /// </summary>
        public static void ClearFonts()
        {
            FontsDictionary.Clear();
            FontCollectionChanged?.Invoke(null, new FontCollectionChangedArg(string.Empty, FontCollectionChangedArg.EventType.Clear));
        }

        /// <summary>
        /// Получение начертания шрифта в виде стороки
        /// </summary>
        /// <param name="bold">Жирный</param>
        /// <param name="italic">Курсив</param>
        /// <returns></returns>
        public static string GetStyleString(bool bold = false, bool italic = false)
        {
            if (bold && italic)
                return @"[BI]";
            if (bold)
                return @"[B]";
            if (italic)
                return @"[I]";
            return @"[R]";
        }

        /// <summary>
        /// Получение семейства шрифта по умолчанию
        /// <returns>Семейство шрифта по умолчанию</returns>
        /// </summary>
        public static string GetDefaultFontFamilyName()
            => GetDefaultFontFamily().FamilyNames.FirstOrDefault().Value ?? string.Empty;

        #endregion

        #region Внутренние поля и методы

        /// <summary>
        /// Информация о событии изменения списка шрифтов
        /// </summary>
        internal class FontCollectionChangedArg : EventArgs
        {

            /// <summary>
            /// Тип события изменения
            /// </summary>
            public enum EventType : ushort
            {
                Delete = 0,
                Clear = 1,
                Add = 2
            }

            /// <summary>
            /// Конструктор
            /// </summary>
            /// <param name="fontName">Имя шрифта</param>
            /// <param name="eventAction">Событие изменения</param>
            public FontCollectionChangedArg(string fontName, EventType eventAction)
            {
                FontName = fontName;
                EventAction = eventAction;
            }
            /// <summary>
            /// Имя шрифта
            /// </summary>
            public string FontName { get; }
            /// <summary>
            /// Тип события изменения
            /// </summary>
            public EventType EventAction { get; }

        }

        /// <summary>
        /// Словарь названия шрифтов и ссылки на файл с шрифтом
        /// </summary>
        internal static Dictionary<string, FontFamily> FontsDictionary { get; set; } = new Dictionary<string, FontFamily>();

        /// <summary>
        /// Событие изменение списка шрифтов
        /// </summary>
        internal static event EventHandler<FontCollectionChangedArg> FontCollectionChanged;

        /// <summary>
        /// Получение FontFamily по имени семейтва и начертанию
        /// </summary>
        /// <param name="familyName">Имя семейства</param>
        /// <param name="bold">Жирный</param>
        /// <param name="italic">Курсив</param>
        /// <returns></returns>
        internal static FontFamily GetFont(string familyName, bool bold = false, bool italic = false)
        {
            var style = GetStyleString(bold, italic);
            var fontName = string.Concat(familyName, style);
            if (FontsDictionary.ContainsKey(fontName))
                return FontsDictionary[fontName];
            var regularFontName = string.Concat(familyName, @"[R]");
            if (FontsDictionary.ContainsKey(regularFontName))
                return FontsDictionary[regularFontName];
            var italicFontName = string.Concat(familyName, @"[I]");
            if (FontsDictionary.ContainsKey(italicFontName))
                return FontsDictionary[italicFontName];
            var boldFontName = string.Concat(familyName, @"[B]");
            if (FontsDictionary.ContainsKey(boldFontName))
                return FontsDictionary[boldFontName];
            var boldItalicFontName = string.Concat(familyName, @"[BI]");
            if (FontsDictionary.ContainsKey(boldItalicFontName))
                return FontsDictionary[boldItalicFontName];
            return GetDefaultFontFamily();
        }

        internal static string GetFontFamilyName(string fontName)
        {
            var regex = new Regex(@"\[[BIR]{1,2}\]$");
            return regex.Replace(fontName, "");
        }

        /// <summary>
        /// Получение шрифта по умолчанию
        /// </summary>
        /// <returns>Шрифт по умолчанию</returns>
        private static FontFamily GetDefaultFontFamily()
            => FontsDictionary.ContainsKey(DefaultFontName)
                ? FontsDictionary[DefaultFontName]
                : FontsDictionary.Any()
                    ? FontsDictionary.First().Value
                    : new FontFamily();

        /// <summary>
        /// Имя семейства шрифта по умолчанию
        /// </summary>
        private static readonly string DefaultFontFamilyName = @"Liberation Sans";

        /// <summary>
        /// Имя шрифта по умолчанию
        /// </summary>
        private static readonly string DefaultFontName = @$"{DefaultFontFamilyName}[R]";

        #endregion

    }
}