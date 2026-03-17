"""
配置管理模块
"""
from pydantic_settings import BaseSettings
from pydantic import Field
from typing import Optional
from enum import Enum


class AIProvider(str, Enum):
    OPENAI = "openai"
    ANTHROPIC = "anthropic"
    DEEPSEEK = "deepseek"
    CUSTOM = "custom"


class Settings(BaseSettings):
    AI_PROVIDER: AIProvider = Field(default=AIProvider.DEEPSEEK, description="AI服务提供商")
    AI_API_KEY: str = Field(default="", description="AI API密钥")
    AI_BASE_URL: Optional[str] = Field(default=None, description="自定义AI API基础URL")
    AI_MODEL: str = Field(default="deepseek-chat", description="AI模型名称")
    
    QUESTIONS_PER_SECTION: int = Field(default=25, description="每节生成的题目数量")
    MAX_TOKENS_PER_REQUEST: int = Field(default=8192, description="每次请求最大token数")
    TEMPERATURE: float = Field(default=0.5, description="生成温度参数")
    
    OUTPUT_DIR: str = Field(default="output", description="输出目录")
    TEMP_DIR: str = Field(default="temp", description="临时文件目录")
    
    PPT_TEMPLATE: Optional[str] = Field(default=None, description="PPT模板路径")
    PPT_TITLE_FONT_SIZE: int = Field(default=36, description="PPT标题字号")
    PPT_CONTENT_FONT_SIZE: int = Field(default=18, description="PPT内容字号")
    
    DOUBAO_API_KEY: str = Field(default="", description="豆包API密钥(用于PPT生成)")
    DOUBAO_BASE_URL: str = Field(default="https://ark.cn-beijing.volces.com/api/v3", description="豆包API基础URL")
    DOUBAO_MODEL: str = Field(default="doubao-lite-32k", description="豆包模型名称")
    DOUBAO_MAX_TOKENS: int = Field(default=30000, description="豆包最大token数")
    DOUBAO_TEMPERATURE: float = Field(default=0.2, description="豆包生成温度")
    
    LOG_LEVEL: str = Field(default="INFO", description="日志级别")
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        case_sensitive = True


settings = Settings()
