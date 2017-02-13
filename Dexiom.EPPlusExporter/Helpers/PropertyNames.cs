using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Dexiom.EPPlusExporter.Helpers
{
    internal static class PropertyNames
    {
        public static string[] For<T, TProp>(Expression<Func<T, TProp>> expression)
        {
            var body = expression.Body;
            return GetMemberNames(body).ToArray();
        }

        public static string[] For<T>(Expression<Func<T, object>> expression)
        {
            var body = expression.Body;
            return GetMemberNames(body).ToArray();
        }

        public static string[] For(Expression<Func<object>> expression)
        {
            var body = expression.Body;
            return GetMemberNames(body).ToArray();
        }
        
        public static IEnumerable<string> GetMemberNames(Expression expression) => GetMemberInfos(expression).Select(n => n.Name);

        public static IEnumerable<MemberInfo> GetMemberInfos(Expression expression)
        {
            if (expression == null || expression is ParameterExpression)
                return Enumerable.Empty<MemberInfo>();

            //most common cases (arguments passed like this: () => MyObject.MyProp1 or n => n.Prop1)
            MemberExpression memberExpression;
            var unary = expression as UnaryExpression;
            if (unary != null)
                memberExpression = unary.Operand as MemberExpression;
            else
                memberExpression = expression as MemberExpression;

            if (memberExpression != null)
                return new[] { memberExpression.Member };

            //arguments passed like this: n => new { n.Prop1, n.Prop2 }
            var newExpression = expression as NewExpression;
            if (newExpression != null)
            {
                return newExpression.Arguments
                    .Select(n => n as MemberExpression)
                    .Where(n => n != null)
                    .Select(n => n.Member);
            }

            return Enumerable.Empty<MemberInfo>();
        }
    }
}