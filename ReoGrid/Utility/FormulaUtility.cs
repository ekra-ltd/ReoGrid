using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using unvell.ReoGrid.Formula;

namespace unvell.ReoGrid.Utility
{
    public sealed class FormulaUtility
    {
        /// <summary>
        /// Состояние разбора
        /// </summary>
        public class State
        {
            public State(string value, int start, int length)
            {
                Value = value;
                Start = start;
                Length = length;
            }

            /// <summary>
            /// Значение
            /// </summary>
            public string Value { get;  }

            /// <summary>
            /// Индекс начала выражения
            /// </summary>
            public int Start { get;  }

            /// <summary>
            /// Длина выражения
            /// </summary>
            public int Length { get; }
        }

        /// <inheritdoc />
        /// <summary>
        /// Состояние разбора. Состояние R1C1 адреса
        /// </summary>
        public class R1C1State: State
        {
            public R1C1State(string value, int start, int length, R1C1Address address): base (value, start, length)
            {
                Address = address;
            }

            /// <summary>
            /// Адрес в формате R1C1
            /// </summary>
            public R1C1Address Address { get; }
        }

        /// <inheritdoc />
        /// <summary>
        /// Состояние разбора. Состояние A1 адреса
        /// </summary>
        public class A1State : State
        {
            public A1State(string value, int start, int length, CellPosition position, string worksheetName) : base(value, start, length)
            {
                Position = position;
                Worksheet = worksheetName;
            }

            /// <summary>
            /// Имя листа, опционально
            /// </summary>
            public string Worksheet { get; }

            /// <summary>
            /// Адрес в формате R1C1
            /// </summary>
            public CellPosition Position { get; }
        }

        /// <summary>
        /// Адрес в формате R1C1
        /// </summary>
        public struct R1C1Address
        {
            /// <summary>
            /// Абсолютный адрес строки
            /// </summary>
            public uint? AbsoluteRow;
            /// <summary>
            /// Абсолютный адрес столбца
            /// </summary>
            public uint? AbsoluteColumn;
            /// <summary>
            /// Относительный адрес строки
            /// </summary>
            public int? RelativeRow;
            /// <summary>
            /// Отсноситльеный адрес столбца
            /// </summary>
            public int? RelativeColumn;
        }

        /// <summary>
        /// Перечиесление идентификаторов R1C1-стиля в формуле
        /// </summary>
        /// <param name="formula">формула</param>
        /// <returns>Список идентификаторов</returns>
        public static IEnumerable<R1C1State> EnumerateR1C1(string formula)
        {
            if (_r1c1Lexer is null)
            {
                _r1c1Lexer = new R1C1Lexer();
            }
            return Enumerate(formula, _r1c1Lexer)
                .Where(t => t.TokenType == TokenType.R1C1Cell)
                .Select(t => t.State as R1C1State)
                .Where(s => s != null);
        }

        /// <summary>
        /// Перечиесление идентификаторов A1-стиля в формуле
        /// </summary>
        /// <param name="formula">формула</param>
        /// <returns>Список идентификаторов</returns>
        public static IEnumerable<A1State> EnumerateA1(string formula)
        {
            if (_a1Lexer is null)
            {
                _a1Lexer = new A1Lexer();
            }

            var tokens = Enumerate(formula, _a1Lexer).ToList();

            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token.TokenType == TokenType.A1Cell)
                {
                    var result = token.State as A1State;
                    if (i >= 2 && result != null)
                    {
                        if (tokens[i - 2].TokenType == TokenType.Identifier &&
                            tokens[i - 1].TokenType == TokenType.Token && tokens[i - 1].State.Value == "!")
                        {
                            result = new A1State(
                                result.Value,
                                tokens[i - 2].State.Start,
                                tokens[i - 2].State.Length + tokens[i - 1].State.Length + tokens[i - 2].State.Length,
                                result.Position,
                                tokens[i - 2].State.Value);

                        }
                    }
                    if (result != null)
                    {
                        yield return result;
                    }
                }
            }
            //return Enumerate(formula, _a1Lexer)
            //    .Where(t => t.TokenType == TokenType.A1Cell)
            //    .Select(t => t.State as A1State)
            //    .Where(s => s != null);
        }

        #region Вспомогательные методы

        private static IEnumerable<Token> Enumerate(string formula, LexerBase lexer)
        {
            if (lexer != null)
            {
                lexer.Reset(formula);
                do
                {
                    var token = lexer.GetToken();
                    if (token != null)
                    {
                        if (token.TokenType == TokenType.Error)
                        {
                            throw new FormulaParseException("Unknown token", lexer.CommittedLength);
                        }
                        yield return token;
                    }
                    else
                    {
                        break;
                    }
                } while (lexer.NextToken());
            }
        }

        #endregion

        #region Поля

        private static LexerBase _r1c1Lexer;
        private static LexerBase _a1Lexer;

        #endregion

        #region Вспомогательные классы

        private class LexerBase
        {
            protected LexerBase(IEnumerable<TokenParser> parsers)
            {
                _parsers = parsers.ToList();

                var regexbuilder = new StringBuilder();
                var first = true;
                foreach (var parser in _parsers)
                {
                    if (first)
                    {
                        first = false;
                    }
                    else
                    {
                        regexbuilder.Append("|");
                    }
                    regexbuilder.Append($"(?<{parser.Name}>{parser.Regex})");
                }
                TokenRegex = new Regex($"\\s*({regexbuilder})", RegexOptions.Compiled);
            }

            public void Reset(string input)
            {
                Start = 0;
                _input = input;
                Length = input.Length;
                Reset();
            }

            // private readonly List<TokenParser> _parsers;

            #region Поля

            private string _input;

            private readonly Regex TokenRegex;

            private List<TokenParser> _parsers;

            private Match _match;

            public int Start { get; set; }
            public int Length { get; set; }
            public int CommittedLength { get; set; }

            #endregion

            #region Методы

            private void Reset()
            {
                CommittedLength = 0;
                _match = TokenRegex.Match(_input, Start);
            }

            public bool NextToken()
            {
                if (_match != null)
                {
                    CommittedLength += _match.Length;

                    if (CommittedLength >= Length)
                    {
                        _match = null;
                    }
                    else
                    {
                        _match = _match.NextMatch();
                    }
                }
                else
                {
                    _match = null;
                }
                return _match != null;
            }

            public  Token GetToken() {
                foreach (var parser in _parsers)
                {
                    var token = parser.TryParse(_match, CommittedLength);
                    if (token != null)
                    {
                        return token;
                    }
                }
                return null;
            }

            #endregion

            #region Вспомогательные классы

           
           
            #endregion

            #region Константы

            #endregion
        }

        private class R1C1Lexer : LexerBase
        {
            public R1C1Lexer(): base(CreateR1C1Parsers()) { }

            private static IEnumerable<TokenParser> CreateR1C1Parsers()
            {
                var decimalSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                return new List<TokenParser>(
                    new[]
                    {
                        new TokenParser(TokenType.String, "string", "\"(?:\"\"|[^\"])*\""),
                        new TokenParser(TokenType.UnionRanges, "union_ranges", "[A-Z]+[0-9]+:[A-Z]+[0-9]+(\\s[A-Z]+[0-9]+:[A-Z]+[0-9]+)+"),
                        new TokenParser(TokenType.Range, "range", "\\$?[A-Z]+\\$?[0-9]*:\\$?[A-Z]+\\$?[0-9]*"),
                        new R1C1TokenParser(),
                        // new TokenParser(TokenType.Cell, "cell", "\\$?[A-Z]+\\$?[0-9]+"),
                        new TokenParser(TokenType.Token, "token", "-"),
                        new TokenParser(TokenType.Number, "number", "\\-?\\d*\\" + decimalSeparator + "?\\d+"),
                        new TokenParser(TokenType.True, "true", "(?i)TRUE"),
                        new TokenParser(TokenType.False, "false", "(?i)FALSE"),
                        new TokenParser(TokenType.Identifier, "identifier", "\\w+"), // Обычный идентификатор
                        new TokenParser(TokenType.Identifier, "identifier", "'([^']|(''))+'"), // Идентификатор, содержащий спецсимволы (", !, ', @, ...)  в имени или начинающиеся с цифры и прочее
                        new TokenParser(TokenType.Token, "token", "\\=\\=|\\<\\>|\\<\\=|\\>\\=|\\<\\>|\\=|\\!|[\\=\\.\\,\\+\\-\\*\\/\\%\\<\\>\\(\\)\\&\\^]"), // скопированы в том числе повторяющиеся токены
                        new TokenParser(TokenType.Error, "error", "."),
                    });
            }
        }

        private class A1Lexer : LexerBase
        {
            public A1Lexer() : base(CreateA1Parsers()) { }

            private static IEnumerable<TokenParser> CreateA1Parsers()
            {
                var decimalSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                return new List<TokenParser>(
                    new[]
                    {
                        new TokenParser(TokenType.String, "string", "\"(?:\"\"|[^\"])*\""),
                        new TokenParser(TokenType.UnionRanges, "union_ranges", "[A-Z]+[0-9]+:[A-Z]+[0-9]+(\\s[A-Z]+[0-9]+:[A-Z]+[0-9]+)+"),
                        new TokenParser(TokenType.Range, "range", "\\$?[A-Z]+\\$?[0-9]*:\\$?[A-Z]+\\$?[0-9]*"),
                        new A1TokenParser(),
                        new TokenParser(TokenType.Token, "token", "-"),
                        new TokenParser(TokenType.Number, "number", "\\-?\\d*\\" + decimalSeparator + "?\\d+"),
                        new TokenParser(TokenType.True, "true", "(?i)TRUE"),
                        new TokenParser(TokenType.False, "false", "(?i)FALSE"),
                        new TokenParser(TokenType.Identifier, "identifier", "\\w+"), // Обычный идентификатор
                        new TokenParser(TokenType.Identifier, "identifier", "'([^']|(''))+'"), // Идентификатор, содержащий спецсимволы (", !, ', @, ...)  в имени или начинающиеся с цифры и прочее
                        new TokenParser(TokenType.Token, "token", "\\=\\=|\\<\\>|\\<\\=|\\>\\=|\\<\\>|\\=|\\!|[\\=\\.\\,\\+\\-\\*\\/\\%\\<\\>\\(\\)\\&\\^]"), // скопированы в том числе повторяющиеся токены
                        new TokenParser(TokenType.Error, "error", "."),
                    });
            }
        }
        #region Парсеры

        enum TokenType
        {
            String,
            UnionRanges,
            Range,
            Cell,
            Token,
            Number,
            True,
            False,
            Identifier,
            Error,
            R1C1Cell,
            A1Cell
        };

        class Token
        {
            public State State { get; set; }

            public TokenType TokenType { get; set; }
        }

        private class TokenParser
        {
            public TokenParser(TokenType type, string name, string regex)
            {
                Type = type;
                Name = name;
                Regex = regex;
            }

            protected TokenType Type { get; }
            public string Name { get; }

            public string Regex { get; }

            public virtual Token TryParse(Match match, int length)
            {
                var group = match.Groups[Name];
                if (group.Success)
                {
                    return new Token
                    {
                        State = new State(group.Value, length, group.Length),
                        TokenType = Type,
                    };
                }
                return null;
            }

        }

        private class R1C1TokenParser : TokenParser
        {
            public R1C1TokenParser() : base(TokenType.R1C1Cell, "r1c1_cell", R1C1RegexPattern) { }

            public override Token TryParse(Match match, int length)
            {
                var group = match.Groups[Name];
                if (group.Success)
                {
                    R1C1Address address = new R1C1Address();
                    if (match.Groups[RelativeRowGroupName].Success)
                    {
                        address.RelativeRow = int.Parse(match.Groups[RelativeRowGroupName].Value);
                    }
                    if (match.Groups[AbsoluteRowGroupName].Success)
                    {
                        address.AbsoluteRow = uint.Parse(match.Groups[AbsoluteRowGroupName].Value);
                    }
                    if (match.Groups[RelativeColumnGroupName].Success)
                    {
                        address.RelativeColumn = int.Parse(match.Groups[RelativeColumnGroupName].Value);
                    }
                    if (match.Groups[AbsoluteColumnGroupName].Success)
                    {
                        address.AbsoluteColumn = uint.Parse(match.Groups[AbsoluteColumnGroupName].Value);
                    }

                    return new Token
                    {
                        State = new R1C1State(group.Value, length, group.Length, address),
                        TokenType = Type,
                    };
                }
                return null;
            }

            private const string RelativeRowGroupName = @"r1c1_RelativeRow";
            private const string AbsoluteRowGroupName = @"r1c1_AbsoluteRow";
            private const string RelativeColumnGroupName = @"r1c1_RelativeColumn";
            private const string AbsoluteColumnGroupName = @"r1c1_AbsoluteColumn";

            private const string R1C1RegexPattern = "R((\\[(?<" + RelativeRowGroupName + ">-?\\d+)\\])|(?<" + AbsoluteRowGroupName + ">\\d+)|())C((\\[(?<" + RelativeColumnGroupName + ">-?\\d+)\\])|(?<" + AbsoluteColumnGroupName + ">\\d+)|())";
        }

        private class A1TokenParser : TokenParser
        {
            public A1TokenParser() : base(TokenType.A1Cell, "a1_cell", A1RegexPattern)
            {
            }

            public override Token TryParse(Match match, int length)
            {
                var group = match.Groups[Name];
                if (group.Success)
                {
                    return new Token
                    {
                        State = new A1State(group.Value, length, group.Length, new CellPosition(group.Value), null),
                        TokenType = Type,
                    };
                }
                return null;
            }

            private const string A1RegexPattern = "\\$?[A-Z]+\\$?[0-9]+";
        }


        #endregion

        #endregion
    }
}
