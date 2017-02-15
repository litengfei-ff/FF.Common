namespace FF.Common.Cache
{
#if NFX461
    using System;
    using System.Runtime.Caching;

    /// <summary>
    /// 基于MemoryCache的缓存辅助类
    /// </summary>
    public class MemoryCacheWriter : ICacheWriter
    {
        public void AddCache(string key, object value)
        {
            var item = new CacheItem(key, value);
            var policy = CreatePolicy(TimeSpan.FromHours(2), null);
            MemoryCache.Default.Add(item, policy);
        }

        public void AddCache(string key, object value, DateTime expDate)
        {
            var item = new CacheItem(key, value);
            var policy = CreatePolicy(null, expDate);
            MemoryCache.Default.Add(item, policy);
        }

        public string GetCache(string key)
        {
            return (string)MemoryCache.Default[key];
        }

        public T GetCache<T>(string key)
        {
            return (T)MemoryCache.Default[key];
        }

        public void RemoveCache(string key)
        {
            MemoryCache.Default.Remove(key);
        }

        public void SetCache(string key, object value)
        {
            if (MemoryCache.Default.Contains(key))
            {
                RemoveCache(key);
            }
            AddCache(key, value);
        }

        public void SetCache(string key, object value, DateTime expDate)
        {
            if (MemoryCache.Default.Contains(key))
            {
                RemoveCache(key);
            }
            AddCache(key, value, expDate);
        }

        /// <summary>
        /// 设置过期信息
        /// </summary>
        /// <param name="slidingExpiration"></param>
        /// <param name="absoluteExpiration"></param>
        /// <returns></returns>
        private static CacheItemPolicy CreatePolicy(TimeSpan? slidingExpiration, DateTime? absoluteExpiration)
        {
            var policy = new CacheItemPolicy();

            if (absoluteExpiration.HasValue)
            {
                policy.AbsoluteExpiration = absoluteExpiration.Value;
            }
            else if (slidingExpiration.HasValue)
            {
                policy.SlidingExpiration = slidingExpiration.Value;
            }

            policy.Priority = CacheItemPriority.Default;

            return policy;
        }
    }
#endif 
}