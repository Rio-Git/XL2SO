using UnityEngine;

namespace XL2SO
{
    /// <summary>
    /// Interface for classes which <see cref="XL2SO"/> runs.
    /// </summary>
    public class IScreen : ScriptableObject
    {
        virtual public void Initialize(XL2SO _parent) { }
        virtual public IScreen Display() { return null; }
    }
}