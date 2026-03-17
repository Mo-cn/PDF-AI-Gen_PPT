"""
AI接口调用模块
"""
import json
from abc import ABC, abstractmethod
from typing import Optional, Dict, Any, List
from openai import OpenAI
from anthropic import Anthropic
import tiktoken

from .config import settings, AIProvider


class BaseAIClient(ABC):
    @abstractmethod
    def generate(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        pass
    
    @abstractmethod
    def count_tokens(self, text: str) -> int:
        pass


class OpenAIClient(BaseAIClient):
    def __init__(self, api_key: str, base_url: Optional[str] = None, model: str = "gpt-4o"):
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model = model
        try:
            self.encoding = tiktoken.encoding_for_model(model)
        except KeyError:
            self.encoding = tiktoken.get_encoding("cl100k_base")
    
    def generate(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                max_tokens=settings.MAX_TOKENS_PER_REQUEST,
                temperature=settings.TEMPERATURE,
                timeout=180.0
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"\nAPI调用错误: {e}")
            raise
    
    def count_tokens(self, text: str) -> int:
        return len(self.encoding.encode(text))


class AnthropicClient(BaseAIClient):
    def __init__(self, api_key: str, model: str = "claude-3-5-sonnet-20241022"):
        self.client = Anthropic(api_key=api_key)
        self.model = model
    
    def generate(self, prompt: str, system_prompt: Optional[str] = None) -> str:
        kwargs = {
            "model": self.model,
            "max_tokens": settings.MAX_TOKENS_PER_REQUEST,
            "messages": [{"role": "user", "content": prompt}]
        }
        if system_prompt:
            kwargs["system"] = system_prompt
        
        response = self.client.messages.create(**kwargs)
        return response.content[0].text
    
    def count_tokens(self, text: str) -> int:
        return len(text) // 4


class DeepSeekClient(OpenAIClient):
    def __init__(self, api_key: str, model: str = "deepseek-chat"):
        super().__init__(
            api_key=api_key,
            base_url="https://api.deepseek.com/v1",
            model=model
        )


class AIClientFactory:
    @staticmethod
    def create_client(
        provider: AIProvider,
        api_key: str,
        base_url: Optional[str] = None,
        model: Optional[str] = None
    ) -> BaseAIClient:
        if provider == AIProvider.OPENAI:
            return OpenAIClient(
                api_key=api_key,
                base_url=base_url,
                model=model or "gpt-4o"
            )
        elif provider == AIProvider.ANTHROPIC:
            return AnthropicClient(
                api_key=api_key,
                model=model or "claude-3-5-sonnet-20241022"
            )
        elif provider == AIProvider.DEEPSEEK:
            return DeepSeekClient(
                api_key=api_key,
                model=model or "deepseek-chat"
            )
        elif provider == AIProvider.CUSTOM:
            if not base_url:
                raise ValueError("自定义AI提供商需要提供base_url")
            return OpenAIClient(
                api_key=api_key,
                base_url=base_url,
                model=model or "gpt-4o"
            )
        else:
            raise ValueError(f"不支持的AI提供商: {provider}")


class AIService:
    def __init__(self):
        self.client = AIClientFactory.create_client(
            provider=settings.AI_PROVIDER,
            api_key=settings.AI_API_KEY,
            base_url=settings.AI_BASE_URL,
            model=settings.AI_MODEL
        )
    
    def analyze_content(self, content: str) -> Dict[str, Any]:
        system_prompt = """你是一个专业的教学内容分析专家。请分析给定的教学内容，提取关键信息。"""
        
        prompt = f"""请分析以下教学内容，并提供以下信息：
1. 核心知识点列表
2. 重点内容摘要
3. 难点分析
4. 建议的教学重点

教学内容：
{content}

请以JSON格式返回结果，格式如下：
{{
    "knowledge_points": ["知识点1", "知识点2", ...],
    "summary": "内容摘要",
    "difficulties": ["难点1", "难点2", ...],
    "teaching_focus": ["重点1", "重点2", ...]
}}"""
        
        response = self.client.generate(prompt, system_prompt)
        
        try:
            json_match = response[response.find('{'):response.rfind('}')+1]
            return json.loads(json_match)
        except:
            return {
                "knowledge_points": [],
                "summary": response,
                "difficulties": [],
                "teaching_focus": []
            }
    
    def generate_questions(
        self,
        content: str,
        num_questions: int = 25,
        question_type: str = "single_choice"
    ) -> List[Dict[str, Any]]:
        system_prompt = """你是一个专业的教育测评专家，擅长设计高质量的选择题。
请确保题目：
1. 准确覆盖核心知识点
2. 难度分布合理（简单30%、中等50%、困难20%）
3. 干扰项具有迷惑性但不误导
4. 题目表述清晰、无歧义"""
        
        prompt = f"""请根据以下教学内容，生成{num_questions}道单选题。

教学内容：
{content}

要求：
1. 每道题包含4个选项（A、B、C、D）
2. 明确标注正确答案
3. 提供简要解析
4. 标注难度级别（easy/medium/hard）
5. 标注涉及的知识点

请以JSON格式返回，格式如下：
{{
    "questions": [
        {{
            "id": "q1",
            "content": "题目内容",
            "options": ["选项A内容", "选项B内容", "选项C内容", "选项D内容"],
            "correct_answer": "A",
            "explanation": "答案解析",
            "difficulty": "medium",
            "knowledge_points": ["知识点1"]
        }},
        ...
    ]
}}"""
        
        response = self.client.generate(prompt, system_prompt)
        
        try:
            json_match = response[response.find('{'):response.rfind('}')+1]
            data = json.loads(json_match)
            return data.get("questions", [])
        except:
            return []
    
    def generate_ppt_content(self, section_title: str, content: str) -> Dict[str, Any]:
        system_prompt = """你是一个专业的PPT内容设计专家，擅长将教学内容转化为清晰的PPT演示文稿。
请确保内容：
1. 结构清晰，逻辑连贯
2. 重点突出，简洁明了
3. 适合课堂教学使用"""
        
        prompt = f"""请将以下教学内容转换为PPT格式的内容结构。

章节标题：{section_title}

教学内容：
{content}

请生成以下内容：
1. 标题页内容
2. 内容页（按知识点分页，每页3-5个要点）
3. 总结页内容

以JSON格式返回：
{{
    "title_slide": {{
        "title": "标题",
        "subtitle": "副标题"
    }},
    "content_slides": [
        {{
            "title": "页面标题",
            "points": ["要点1", "要点2", "要点3"]
        }}
    ],
    "summary_slide": {{
        "title": "本章总结",
        "points": ["总结要点1", "总结要点2", ...]
    }}
}}"""
        
        response = self.client.generate(prompt, system_prompt)
        
        try:
            json_match = response[response.find('{'):response.rfind('}')+1]
            return json.loads(json_match)
        except:
            return {
                "title_slide": {"title": section_title, "subtitle": ""},
                "content_slides": [],
                "summary_slide": {"title": "本章总结", "points": []}
            }
