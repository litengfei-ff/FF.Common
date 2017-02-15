using System;

namespace FF.Common.Cache
{
    public class CacheHelper
    {

        private static ICacheWriter CacheWriter { get; set; }

        static CacheHelper()
        {
#if NFX461
            CacheWriter = new MemoryCacheWriter();
#elif NETCORE
            CacheWriter = null;
#endif
        }


        public static void AddCache(string key, object value, DateTime expDate)
        {
            CacheWriter.AddCache(key, value, expDate);
        }


        public static void AddCache(string key, object value)
        {
            CacheWriter.AddCache(key, value);
        }

        public static T GetCache<T>(string key)
        {
            return CacheWriter.GetCache<T>(key);
        }

        public static string GetCache(string key)
        {
            return CacheWriter.GetCache(key);
        }

        public static void SetCache(string key, object value, DateTime extTime)
        {
            CacheWriter.SetCache(key, value, extTime);
        }

        public static void SetCache(string key, object value)
        {
            CacheWriter.SetCache(key, value);
        }

        public static void RemoveCache(string key)
        {
            CacheWriter.RemoveCache(key);
        }
    }
}
