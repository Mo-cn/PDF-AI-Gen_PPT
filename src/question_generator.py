"""
试题自动生成模块
"""
import uuid
import signal
import sys
from typing import List, Dict, Any
from tqdm import tqdm

from .models import Question, QuestionSet, Section, QuestionType, Difficulty
from .ai_client import AIService
from .config import settings

_interrupted = False

def signal_handler(signum, frame):
    global _interrupted
    print("\n\n收到中断信号，正在保存已生成的内容...")
    _interrupted = True

signal.signal(signal.SIGINT, signal_handler)


class QuestionGenerator:
    def __init__(self):
        self.ai_service = AIService()
    
    def generate_for_section(
        self,
        section: Section,
        num_questions: int = None,
        parent_title: str = None
    ) -> QuestionSet:
        num_questions = num_questions or settings.QUESTIONS_PER_SECTION
        
        raw_questions = self.ai_service.generate_questions(
            content=section.content,
            num_questions=num_questions
        )
        
        questions = []
        for i, q in enumerate(raw_questions):
            try:
                options = q.get("options", [])
                if not isinstance(options, list) or len(options) < 4:
                    continue
                
                options = [str(opt).strip() if opt else "" for opt in options[:4]]
                while len(options) < 4:
                    options.append("")
                
                if not all(options):
                    continue
                
                content = q.get("content", "")
                if not content or len(content) < 5:
                    continue
                
                correct_answer = str(q.get("correct_answer", "A")).upper()
                if correct_answer not in ["A", "B", "C", "D"]:
                    correct_answer = "A"
                
                question = Question(
                    id=f"{section.id}_q{i+1}",
                    section_id=section.id,
                    question_type=QuestionType.SINGLE_CHOICE,
                    difficulty=self._parse_difficulty(q.get("difficulty", "medium")),
                    content=content,
                    options=options,
                    correct_answer=correct_answer,
                    explanation=q.get("explanation"),
                    knowledge_points=q.get("knowledge_points", [])
                )
                questions.append(question)
            except Exception as e:
                print(f"    跳过无效题目: {str(e)[:30]}")
                continue
        
        return QuestionSet(
            section_id=section.id,
            section_title=section.title,
            parent_title=parent_title,
            questions=questions,
            total_count=len(questions)
        )
    
    def _parse_difficulty(self, difficulty_str: str) -> Difficulty:
        difficulty_map = {
            "easy": Difficulty.EASY,
            "simple": Difficulty.EASY,
            "简单": Difficulty.EASY,
            "medium": Difficulty.MEDIUM,
            "normal": Difficulty.MEDIUM,
            "中等": Difficulty.MEDIUM,
            "hard": Difficulty.HARD,
            "difficult": Difficulty.HARD,
            "困难": Difficulty.HARD,
        }
        return difficulty_map.get(difficulty_str.lower(), Difficulty.MEDIUM)
    
    def generate_for_all_sections(
        self,
        sections: List[Section],
        num_questions_per_section: int = None,
        min_content_length: int = 100
    ) -> List[QuestionSet]:
        global _interrupted
        question_sets = []
        
        flat_sections = self._flatten_sections(sections)
        valid_sections = [(s, p) for s, p in flat_sections if len(s.content.strip()) >= min_content_length]
        
        for i, (section, parent_title) in enumerate(valid_sections, 1):
            if _interrupted:
                break
            
            try:
                question_set = self.generate_for_section(
                    section=section,
                    num_questions=num_questions_per_section,
                    parent_title=parent_title
                )
                question_sets.append(question_set)
            except KeyboardInterrupt:
                _interrupted = True
            except Exception:
                continue
        
        return question_sets
    
    def _flatten_sections(self, sections: List[Section], parent_title: str = None) -> List[tuple]:
        result = []
        for section in sections:
            result.append((section, parent_title))
            if section.children:
                result.extend(self._flatten_sections(section.children, section.title))
        return result
    
    def validate_questions(self, question_set: QuestionSet) -> Dict[str, Any]:
        validation_result = {
            "is_valid": True,
            "errors": [],
            "warnings": [],
            "statistics": {
                "total": question_set.total_count,
                "by_difficulty": {
                    "easy": 0,
                    "medium": 0,
                    "hard": 0
                }
            }
        }
        
        for question in question_set.questions:
            validation_result["statistics"]["by_difficulty"][question.difficulty.value] += 1
            
            if len(question.options) != 4:
                validation_result["errors"].append(
                    f"题目 {question.id} 选项数量不正确: {len(question.options)}"
                )
                validation_result["is_valid"] = False
            
            if question.correct_answer not in ["A", "B", "C", "D"]:
                validation_result["errors"].append(
                    f"题目 {question.id} 正确答案格式不正确: {question.correct_answer}"
                )
                validation_result["is_valid"] = False
            
            if not question.content.strip():
                validation_result["warnings"].append(
                    f"题目 {question.id} 内容为空"
                )
            
            for i, option in enumerate(question.options):
                if not option.strip():
                    validation_result["warnings"].append(
                        f"题目 {question.id} 选项 {chr(65+i)} 为空"
                    )
        
        total = validation_result["statistics"]["total"]
        if total > 0:
            easy_ratio = validation_result["statistics"]["by_difficulty"]["easy"] / total
            medium_ratio = validation_result["statistics"]["by_difficulty"]["medium"] / total
            hard_ratio = validation_result["statistics"]["by_difficulty"]["hard"] / total
            
            if easy_ratio < 0.2 or easy_ratio > 0.4:
                validation_result["warnings"].append(
                    f"简单题比例 {easy_ratio:.1%} 不在建议范围 20%-40% 内"
                )
            if medium_ratio < 0.4 or medium_ratio > 0.6:
                validation_result["warnings"].append(
                    f"中等题比例 {medium_ratio:.1%} 不在建议范围 40%-60% 内"
                )
            if hard_ratio < 0.1 or hard_ratio > 0.3:
                validation_result["warnings"].append(
                    f"困难题比例 {hard_ratio:.1%} 不在建议范围 10%-30% 内"
                )
        
        return validation_result
    
    def filter_questions(
        self,
        question_set: QuestionSet,
        min_count: int = 20,
        remove_duplicates: bool = True
    ) -> QuestionSet:
        filtered_questions = question_set.questions
        
        if remove_duplicates:
            seen_contents = set()
            unique_questions = []
            for q in filtered_questions:
                content_key = q.content.strip().lower()
                if content_key not in seen_contents:
                    seen_contents.add(content_key)
                    unique_questions.append(q)
            filtered_questions = unique_questions
        
        if len(filtered_questions) < min_count:
            pass
        else:
            difficulty_order = {
                Difficulty.EASY: 0,
                Difficulty.MEDIUM: 1,
                Difficulty.HARD: 2
            }
            filtered_questions.sort(
                key=lambda q: (difficulty_order[q.difficulty], q.id)
            )
        
        return QuestionSet(
            section_id=question_set.section_id,
            section_title=question_set.section_title,
            questions=filtered_questions,
            total_count=len(filtered_questions)
        )
