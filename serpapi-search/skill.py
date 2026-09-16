import os
from serpapi import Client  # 注意：0.1.5 版本用 Client 而不是 GoogleSearch

class SerpApiSkill:
    def __init__(self, api_key=None):
        # 优先使用传入的密钥，否则从环境变量读取
        self.api_key = api_key or os.getenv('SERPAPI_API_KEY')
        if not self.api_key:
            raise ValueError("请设置 SERPAPI_API_KEY 环境变量或在初始化时传入")
        
        # 初始化客户端
        self.client = Client(api_key=self.api_key)
        print(f"✅ SerpApi 初始化成功")
    
    def search(self, query: str, num_results: int = 5) -> str:
        """
        使用 SerpApi 进行网页搜索
        
        Args:
            query: 搜索关键词
            num_results: 返回结果数量
        
        Returns:
            格式化的搜索结果字符串
        """
        print(f"🔍 正在搜索: {query}")
        
        try:
            # 执行搜索（0.1.5 版本的用法）
            results = self.client.search(q=query, num=num_results)
            
            # 打印返回结果的结构（调试用）
            print(f"📊 返回结果类型: {type(results)}")
            
            # 尝试不同的方式获取结果
            if hasattr(results, 'get'):
                # 如果是字典
                organic_results = results.get('organic_results', [])
            elif hasattr(results, 'organic_results'):
                # 如果是对象
                organic_results = results.organic_results
            else:
                # 尝试转换为字典
                try:
                    results_dict = results.__dict__
                    organic_results = results_dict.get('organic_results', [])
                except:
                    organic_results = []
            
            if not organic_results:
                # 尝试获取前几个结果
                if hasattr(results, 'search_metadata'):
                    return f"✅ 搜索完成，但没有找到结果。原始返回：{str(results)[:200]}"
                else:
                    return "❌ 没有找到相关结果"
            
            # 格式化输出
            output = []
            output.append(f"📊 找到 {len(organic_results)} 条结果：\n")
            
            for i, result in enumerate(organic_results[:num_results], 1):
                # 兼容不同的属性访问方式
                if hasattr(result, 'title'):
                    title = result.title
                    link = getattr(result, 'link', getattr(result, 'url', '无链接'))
                    snippet = getattr(result, 'snippet', getattr(result, 'description', '无摘要'))
                else:
                    title = result.get('title', '无标题')
                    link = result.get('link', result.get('url', '无链接'))
                    snippet = result.get('snippet', result.get('description', '无摘要'))
                
                output.append(f"{i}. {title}")
                output.append(f"   📎 {link}")
                output.append(f"   📝 {snippet}")
                output.append("")
            
            return "\n".join(output)
        
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(f"❌ 错误详情: {error_detail}")
            return f"❌ 搜索出错: {str(e)}"
    
    def search_news(self, query: str, num_results: int = 5) -> str:
        """专门搜索新闻"""
        try:
            # 新闻搜索参数
            results = self.client.search(q=query, num=num_results, tbm='nws')
            
            # 获取新闻结果
            if hasattr(results, 'news_results'):
                news_results = results.news_results
            elif hasattr(results, 'get'):
                news_results = results.get('news_results', [])
            else:
                news_results = []
            
            if not news_results:
                return "❌ 没有找到相关新闻"
            
            output = [f"📰 找到 {len(news_results)} 条新闻：\n"]
            for i, news in enumerate(news_results[:num_results], 1):
                if hasattr(news, 'title'):
                    title = news.title
                    link = getattr(news, 'link', getattr(news, 'url', '无链接'))
                    source = getattr(news, 'source', '未知来源')
                    date = getattr(news, 'date', '未知日期')
                else:
                    title = news.get('title', '无标题')
                    link = news.get('link', news.get('url', '无链接'))
                    source = news.get('source', '未知来源')
                    date = news.get('date', '未知日期')
                
                output.append(f"{i}. {title}")
                output.append(f"   📎 {link}")
                output.append(f"   🏢 {source} | 📅 {date}")
                output.append("")
            
            return "\n".join(output)
        
        except Exception as e:
            return f"❌ 新闻搜索出错: {str(e)}"

# CoPaw 调用的入口函数
def websearch(query: str) -> str:
    """CoPaw 调用的入口函数"""
    api_key = os.getenv('SERPAPI_API_KEY')
    if not api_key:
        return "❌ 请先设置 SERPAPI_API_KEY 环境变量"
    
    skill = SerpApiSkill(api_key)
    return skill.search(query)

# 测试代码
if __name__ == "__main__":
    # 从环境变量获取 API 密钥
    api_key = os.getenv('SERPAPI_API_KEY')
    if not api_key:
        print("❌ 请先设置 SERPAPI_API_KEY 环境变量")
        print("例如: $env:SERPAPI_API_KEY='你的密钥'")
    else:
        # 测试搜索
        skill = SerpApiSkill(api_key)
        result = skill.search("Python 教程")
        print(result)