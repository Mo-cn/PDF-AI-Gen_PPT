"""
输出管理模块
"""
import json
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime
import pandas as pd

from .models import QuestionSet, Question, PDFDocument, PPTDocument
from .config import settings


class OutputManager:
    def __init__(self, output_dir: str = None):
        self.output_dir = Path(output_dir or settings.OUTPUT_DIR)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def save_questions_to_json(
        self,
        question_sets: List[QuestionSet],
        filename: str = None
    ) -> str:
        filename = filename or f"questions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        file_path = self.output_dir / filename
        
        data = {
            "generated_at": datetime.now().isoformat(),
            "total_sections": len(question_sets),
            "total_questions": sum(qs.total_count for qs in question_sets),
            "sections": []
        }
        
        for qs in question_sets:
            section_data = {
                "section_id": qs.section_id,
                "section_title": qs.section_title,
                "question_count": qs.total_count,
                "questions": [q.model_dump() for q in qs.questions]
            }
            data["sections"].append(section_data)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return str(file_path)
    
    def save_questions_to_excel(
        self,
        question_sets: List[QuestionSet],
        filename: str = None,
        separate_answer: bool = True
    ) -> Dict[str, str]:
        filename = filename or f"questions_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        all_questions = []
        for qs in question_sets:
            for q in qs.questions:
                question_data = {
                    "章节ID": q.section_id,
                    "章节标题": qs.section_title,
                    "题目ID": q.id,
                    "题目内容": q.content,
                    "选项A": q.options[0] if len(q.options) > 0 else "",
                    "选项B": q.options[1] if len(q.options) > 1 else "",
                    "选项C": q.options[2] if len(q.options) > 2 else "",
                    "选项D": q.options[3] if len(q.options) > 3 else "",
                    "正确答案": q.correct_answer,
                    "答案解析": q.explanation or "",
                    "难度": q.difficulty.value,
                    "知识点": ", ".join(q.knowledge_points)
                }
                all_questions.append(question_data)
        
        df = pd.DataFrame(all_questions)
        
        result = {}
        
        if separate_answer:
            questions_df = df.drop(columns=["正确答案", "答案解析"])
            questions_path = self.output_dir / f"{filename}_题目.xlsx"
            questions_df.to_excel(questions_path, index=False, engine='openpyxl')
            result["questions"] = str(questions_path)
            
            answers_df = df[["题目ID", "正确答案", "答案解析"]]
            answers_path = self.output_dir / f"{filename}_答案.xlsx"
            answers_df.to_excel(answers_path, index=False, engine='openpyxl')
            result["answers"] = str(answers_path)
        else:
            combined_path = self.output_dir / f"{filename}.xlsx"
            df.to_excel(combined_path, index=False, engine='openpyxl')
            result["combined"] = str(combined_path)
        
        return result
    
    def save_document_structure(
        self,
        document: PDFDocument,
        filename: str = None
    ) -> str:
        filename = filename or f"structure_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        file_path = self.output_dir / filename
        
        def section_to_dict(section):
            return {
                "id": section.id,
                "title": section.title,
                "level": section.level,
                "content_length": len(section.content),
                "page_range": f"{section.page_start}-{section.page_end}" if section.page_start else None,
                "children": [section_to_dict(child) for child in section.children]
            }
        
        data = {
            "file_path": document.file_path,
            "title": document.title,
            "total_pages": document.total_pages,
            "metadata": document.metadata,
            "sections": [section_to_dict(s) for s in document.sections],
            "generated_at": datetime.now().isoformat()
        }
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return str(file_path)
    
    def save_ppt_manifest(
        self,
        ppt_documents: List[PPTDocument],
        filename: str = None
    ) -> str:
        filename = filename or f"ppt_manifest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        file_path = self.output_dir / filename
        
        data = {
            "generated_at": datetime.now().isoformat(),
            "total_files": len(ppt_documents),
            "files": [
                {
                    "section_id": doc.section_id,
                    "section_title": doc.section_title,
                    "output_path": doc.output_path,
                    "slide_count": len(doc.slides)
                }
                for doc in ppt_documents
            ]
        }
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return str(file_path)
    
    def generate_report(
        self,
        document: PDFDocument,
        question_sets: List[QuestionSet],
        ppt_documents: List[PPTDocument],
        filename: str = None
    ) -> str:
        filename = filename or f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
        file_path = self.output_dir / filename
        
        total_questions = sum(qs.total_count for qs in question_sets)
        
        difficulty_stats = {"easy": 0, "medium": 0, "hard": 0}
        for qs in question_sets:
            for q in qs.questions:
                difficulty_stats[q.difficulty.value] += 1
        
        report = f"""# PDF内容处理与PPT自动生成报告

## 基本信息
- **PDF文件**: {document.file_path}
- **文档标题**: {document.title}
- **总页数**: {document.total_pages}
- **生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 章节结构
- **章节数量**: {len(document.sections)}

### 章节列表
"""
        
        def add_sections(sections, level=0):
            nonlocal report
            for section in sections:
                indent = "  " * level
                report += f"{indent}- {section.title} (层级: {section.level})\n"
                if section.children:
                    add_sections(section.children, level + 1)
        
        add_sections(document.sections)
        
        report += f"""
## 试题生成统计
- **总题目数**: {total_questions}
- **章节数**: {len(question_sets)}

### 难度分布
- 简单题: {difficulty_stats['easy']} ({difficulty_stats['easy']/total_questions*100:.1f}%)
- 中等题: {difficulty_stats['medium']} ({difficulty_stats['medium']/total_questions*100:.1f}%)
- 困难题: {difficulty_stats['hard']} ({difficulty_stats['hard']/total_questions*100:.1f}%)

### 各章节题目数量
"""
        
        for qs in question_sets:
            report += f"- {qs.section_title}: {qs.total_count} 题\n"
        
        report += f"""
## PPT生成统计
- **PPT文件数**: {len(ppt_documents)}

### 生成的PPT文件
"""
        
        for doc in ppt_documents:
            report += f"- [{doc.section_title}]({doc.output_path})\n"
        
        report += """
---
*本报告由 PDF-AI-GEN_PPT 系统自动生成*
"""
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        return str(file_path)
    
    def load_questions_from_json(self, file_path: str) -> List[QuestionSet]:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        question_sets = []
        for section_data in data.get("sections", []):
            questions = []
            for q_data in section_data.get("questions", []):
                question = Question(**q_data)
                questions.append(question)
            
            question_set = QuestionSet(
                section_id=section_data["section_id"],
                section_title=section_data["section_title"],
                questions=questions,
                total_count=len(questions)
            )
            question_sets.append(question_set)
        
        return question_sets
