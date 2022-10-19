using System;
using Antlr4.Runtime;

namespace Application
{
    public class ThrowExceptionErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        private const string IncorrectExpression = "Некоректний вираз: {0}";
        //BaseErrorListener implementation
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            throw new ArgumentException(IncorrectExpression, msg, e);
        }

        //IAntlrErrorListener<int> implementation
        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            throw new ArgumentException(IncorrectExpression, msg, e);
        }
    }
}
