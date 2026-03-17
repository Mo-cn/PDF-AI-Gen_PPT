"""
数据模型定义
"""
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum


class QuestionType(str, Enum):
    SINGLE_CHOICE = "single_choice"
    MULTIPLE_CHOICE = "multiple_choice"


class Difficulty(str, Enum):
    EASY = "easy"
    MEDIUM = "medium"
    HARD = "hard"


class Question(BaseModel):
    id: str = Field(description="题目唯一标识")
    section_id: str = Field(description="所属章节ID")
    question_type: QuestionType = Field(default=QuestionType.SINGLE_CHOICE, description="题目类型")
    difficulty: Difficulty = Field(default=Difficulty.MEDIUM, description="难度级别")
    content: str = Field(description="题目内容")
    options: List[str] = Field(description="选项列表")
    correct_answer: str = Field(description="正确答案（选项字母）")
    explanation: Optional[str] = Field(default=None, description="答案解析")
    knowledge_points: List[str] = Field(default_factory=list, description="涉及的知识点")


class Section(BaseModel):
    id: str = Field(description="章节唯一标识")
    title: str = Field(description="章节标题")
    level: int = Field(description="章节层级")
    content: str = Field(description="章节内容")
    parent_id: Optional[str] = Field(default=None, description="父章节ID")
    children: List["Section"] = Field(default_factory=list, description="子章节列表")
    page_start: Optional[int] = Field(default=None, description="起始页码")
    page_end: Optional[int] = Field(default=None, description="结束页码")


class PDFDocument(BaseModel):
    file_path: str = Field(description="PDF文件路径")
    title: str = Field(description="文档标题")
    total_pages: int = Field(description="总页数")
    sections: List[Section] = Field(default_factory=list, description="章节列表")
    metadata: Dict[str, Any] = Field(default_factory=dict, description="元数据")


class QuestionSet(BaseModel):
    section_id: str = Field(description="章节ID")
    section_title: str = Field(description="章节标题")
    questions: List[Question] = Field(description="题目列表")
    total_count: int = Field(description="题目总数")


class PPTSlide(BaseModel):
    slide_type: str = Field(description="幻灯片类型：title, content, summary, question")
    title: str = Field(description="幻灯片标题")
    content: List[str] = Field(default_factory=list, description="内容列表")
    notes: Optional[str] = Field(default=None, description="演讲者备注")


class PPTDocument(BaseModel):
    section_id: str = Field(description="章节ID")
    section_title: str = Field(description="章节标题")
    slides: List[PPTSlide] = Field(default_factory=list, description="幻灯片列表")
    output_path: Optional[str] = Field(default=None, description="输出路径")


Section.model_rebuild()
