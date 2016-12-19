namespace DllActivator
{ 
    /// <summary> 用于 OldW Revit 插件中多个dll之前的AddinManager调试 </summary>
    public class DllActivator_eZx_API : IDllActivator_std
    {
        /// <summary>
        /// 激活本DLL所引用的那些DLLs
        /// </summary>
        public void ActivateReferences()
        {
            DllActivator_std dat1 = new DllActivator_std();
            dat1.ActivateReferences();
        }
    }
}