from collections import Counter
import numpy as np
import requests
import json
import time

def calculate_cosine_similarity(str1, str2):
    """
    计算两个字符串的余弦相似度。

    参数:
    - str1: 第一个字符串，用于与第二个字符串比较相似度。
    - str2: 第二个字符串，用于与第一个字符串比较相似度。
    
    返回:
    - 一个浮点数，表示两个字符串的余弦相似度。
      当两个字符串完全相同时，返回1；当两个字符串没有共同的字符时，返回0。
    
    异常:
    - 如果任一输入不是字符串类型，则抛出ValueError异常。
    """
    
    # 检查输入类型
    if not isinstance(str1, str) or not isinstance(str2, str):
        raise ValueError("输入参数必须为字符串类型")
    
    # 统计字符出现频次
    co_str1 = Counter(str1)
    co_str2 = Counter(str2)
    
    # 获取所有独特的字符
    all_chars = set(str1).union(set(str2))
    
    # 构建向量
    vec_str1 = np.array([co_str1.get(char, 0) for char in all_chars])
    vec_str2 = np.array([co_str2.get(char, 0) for char in all_chars])
    
    # 防止除以零错误
    if np.linalg.norm(vec_str1) == 0 or np.linalg.norm(vec_str2) == 0:
        return 0.0
    
    # 计算余弦相似度
    return np.dot(vec_str1, vec_str2) / (np.linalg.norm(vec_str1) * np.linalg.norm(vec_str2))

def find_most_similar(target_string, string_list):
    """
    在给定的字符串列表中找出与目标字符串最相似的元素及其相似度。
    
    参数:
    - target_string: 用于比较的基准字符串。
    - string_list: 包含多个字符串的列表，用于与target_string进行比较。
    
    返回:
    - 一个元组，包含两个元素：
      - 最相似的字符串；
      - 该字符串与target_string的余弦相似度。
    
    注意:
    - 如果列表为空或所有字符串与target_string的相似度为0，则返回None和0。
    """
    
    max_similarity = -1
    most_similar_element = None
    for element in string_list:
        similarity = calculate_cosine_similarity(target_string, element)
        if similarity > max_similarity:
            max_similarity = similarity
            most_similar_element = element
            
    return most_similar_element, max_similarity


def extract_chinese_characters(s):
    """
    提取字符串中的所有中文字符。
    
    参数:
    - s: 输入的字符串。
    
    返回:
    - 一个新的字符串，仅包含输入字符串中的中文字符。
    """
    # 正则表达式匹配中文字符
    pattern = r'[\u4e00-\u9fff]'
    # 使用 re.findall 查找所有匹配的中文字符
    chinese_chars = re.findall(pattern, s)
    # 将所有匹配的中文字符连接成一个字符串
    return ''.join(chinese_chars)



def get_city(address, max_retries=3, delay=5):
    """
    从给定的地址中获取省、市、区的行政区划信息。
    
    参数:
    - address: 地址字符串。
    - max_retries: 请求失败时的最大重试次数，默认为3次。
    - delay: 两次请求之间的延迟时间（秒），默认为5秒。
    
    返回:
    - 一个元组，包含三个元素：省份、城市、区县。
      如果数据无法获取或地址无效，返回值可能包含None。
    
    注意:
    - 需要一个有效的高德地图API密钥才能使用此功能。
    - API请求可能会受到限制或产生费用，请参考高德地图的官方文档了解详情。
    """
    
    # 检查输入类型
    if not isinstance(address, str):
        return None, None, None
    
    # 清洗地址
    address = address.split('#')[0].strip()
    if not address:
        return None, None, None
    
    # 构造请求URL
    api_key = ''  # 这里应使用你的高德地图API密钥
    url = f'https://restapi.amap.com/v3/geocode/geo?address={address}&output=JSON&key={api_key}'
    
    # 发送请求并处理响应
    for attempt in range(max_retries):
        try:
            response = requests.get(url)
            response.raise_for_status()  # 检查HTTP状态码是否为200
            result = response.json()
            
            # 解析结果
            if 'geocodes' in result and result['geocodes']:
                geocode_info = result['geocodes'][0]
                province = geocode_info.get('province', None)
                city = geocode_info.get('city', None)
                district = geocode_info.get('district', None)
                
                return province, city, district
            
            # 如果没有找到地理编码信息
            return None, None, None
        
        except (requests.exceptions.RequestException, KeyError) as e:
            print(f"Error fetching data for address '{address}': {e}")
            
            # 检查是否达到最大重试次数
            if attempt < max_retries - 1:
                print(f"Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                print(f"Failed to fetch data for address '{address}' after {max_retries} attempts")
                return None, None, None
