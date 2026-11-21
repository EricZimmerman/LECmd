#if NETFRAMEWORK
namespace System.Runtime.CompilerServices
{
    /// <summary>CsvHelper 34's generated code registers its mappers with a module initializer, which .NET Framework lacks</summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    internal sealed class ModuleInitializerAttribute : Attribute
    {
    }
}
#endif
