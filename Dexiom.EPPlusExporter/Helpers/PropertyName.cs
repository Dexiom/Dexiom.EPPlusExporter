using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter.Helpers
{
    internal static class PropertyName
    {
        public static string For<T, TProp>(Expression<Func<T, TProp>> expression)
        {
            var body = expression.Body;
            return GetMemberName(body);
        }
        public static string For<T>(Expression<Func<T, object>> expression)
        {
            var body = expression.Body;
            return GetMemberName(body);
        }
        public static string For(Expression<Func<object>> expression)
        {
            var body = expression.Body;
            return GetMemberName(body);
        }
        public static string GetMemberName(Expression expression)
        {
            MemberExpression memberExpression;

            var unary = expression as UnaryExpression;
            if (unary != null)
                //In this case the return type of the property was not object,
                //so .Net wrapped the expression inside of a unary Convert()
                //expression that casts it to type object. In this case, the
                //Operand of the Convert expression has the original expression.
                memberExpression = unary.Operand as MemberExpression;
            else
                //when the property is of type object the body itself is the
                //correct expression
                memberExpression = expression as MemberExpression;

            if (memberExpression == null)
                throw new ArgumentException(
                    "Expression was not of the form 'x => x.Property or x => x.Field'.");

            return memberExpression.Member.Name;
        }
    }
}
