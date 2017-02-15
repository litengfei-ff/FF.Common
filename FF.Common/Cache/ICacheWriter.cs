using System;

namespace FF.Common.Cache
{
    
    /// <summary>
    /// 缓存接口
    /// </summary>
    public interface ICacheWriter
    {
        void AddCache(string key, object value, DateTime expDate);
        void AddCache(string key, object value);
        T GetCache<T>(string key);
        string GetCache(string key);
        void SetCache(string key, object value, DateTime expDate);
        void SetCache(string key, object value);
        void RemoveCache(string key);
    }
}