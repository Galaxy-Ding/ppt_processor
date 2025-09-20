import re
from typing import Optional

def split_after_colon(
    input_str: str,
    clean_whitespace: bool = False,
    recursive: bool = True
) -> Optional[str]:
    """
    递归分割字符串中的中英文冒号（:或：）后的内容
    
    Args:
        input_str: 待分割的输入字符串
        clean_whitespace: 是否清理分割结果首尾的空白
        recursive: 是否递归分割剩余冒号
        
    Returns:
        Optional[str]: 分割后的字符串，若输入无效则返回None
    """
    if not isinstance(input_str, str):
        return None

    pattern = r'\s*[:\uFF1A]\s*'
    parts = re.split(pattern, input_str, maxsplit=1)

    if len(parts) < 2:
        return input_str.strip() if clean_whitespace else input_str

    _, after = parts
    if clean_whitespace:
        after = after.strip()

    if recursive and any(c in after for c in ':：'):
        return split_after_colon(after, clean_whitespace, recursive)

    return after.strip() if clean_whitespace else after

def extract_text_by_pattern(
    text: str, 
    pattern: str, 
    ignore_case: bool = False
) -> str:
    """
    使用正则表达式提取文本中的特定内容
    
    Args:
        text: 源文本
        pattern: 正则表达式模式
        ignore_case: 是否忽略大小写
        
    Returns:
        str: 匹配到的文本，未匹配则返回空字符串
    """
    regex = re.compile(pattern, flags=re.IGNORECASE if ignore_case else 0)
    match = regex.search(text)
    
    if match and len(match.regs) > 0:
        matched_text = text[match.regs[0][0]:match.regs[0][1]]
        return split_after_colon(matched_text) or ""
    
    return ""

# 常用的正则表达式模式
COMMON_PATTERNS = {
    "数字百分比": r"\d+(\.\d+)?%",
    "比较运算": r"[≥≤=><]\s*\d+(\.\d+)?%",
    "中英文冒号": r"[:：]",
    "空白字符": r"\s+",
}

if __name__ == "__main__":
    # 使用示例
    text = "测试数据：123.45%"
    result = split_after_colon(text)
    print(f"分割结果: {result}")  # 输出: 123.45%
    
    pattern = r"数据[:：]\s*(\d+\.\d+%)"
    extracted = extract_text_by_pattern(text, pattern)
    print(f"提取结果: {extracted}")  # 输出: 123.45%

    # 分割冒号后的内容
    text = "标题：内容"
    content = split_after_colon(text)

    # 使用正则提取内容
    pattern = r"标题[:：]\s*(.*)"
    extracted = extract_text_by_pattern(text, pattern)