"""
PDF解析模块
"""
import re
import uuid
from pathlib import Path
from typing import List, Optional, Tuple, Dict
import pdfplumber
from PyPDF2 import PdfReader
from tqdm import tqdm

from .models import PDFDocument, Section
from .config import settings


class PDFParser:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"PDF文件不存在: {file_path}")
        
        self.pdf_plumber = None
        self.pdf_reader = None
        self._open_files()
    
    def _open_files(self):
        self.pdf_plumber = pdfplumber.open(self.file_path)
        self.pdf_reader = PdfReader(str(self.file_path))
    
    def close(self):
        if self.pdf_plumber:
            self.pdf_plumber.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def get_metadata(self) -> dict:
        metadata = {}
        if self.pdf_reader.metadata:
            metadata = {
                "title": self.pdf_reader.metadata.get("/Title", ""),
                "author": self.pdf_reader.metadata.get("/Author", ""),
                "subject": self.pdf_reader.metadata.get("/Subject", ""),
                "creator": self.pdf_reader.metadata.get("/Creator", ""),
                "producer": self.pdf_reader.metadata.get("/Producer", ""),
                "creation_date": str(self.pdf_reader.metadata.get("/CreationDate", "")),
            }
        return metadata
    
    def get_outlines(self) -> List[Dict]:
        """获取PDF书签/大纲结构"""
        try:
            outlines = self.pdf_reader.outline
            if not outlines:
                return []
            
            result = []
            self._parse_outline_recursive(outlines, result, level=0)
            
            for i, item in enumerate(result):
                print(f"    书签 {i+1}: 第{item['page']}页 [层级{item['level']}] {item['title'][:40]}")
            
            return result
        except Exception as e:
            print(f"  解析书签出错: {e}")
            return []
    
    def _parse_outline_recursive(self, items: List, result: List, level: int):
        """递归解析书签结构
        
        PyPDF2的outline结构：
        [dict1, [子书签列表], dict2, [子书签列表], ...]
        """
        i = 0
        while i < len(items):
            item = items[i]
            
            if isinstance(item, dict):
                try:
                    title = item.get('/Title', '')
                    page_num = 1
                    try:
                        page_num = self.pdf_reader.get_destination_page_number(item) + 1
                    except:
                        if '/A' in item:
                            dest = item['/A']
                            if '/D' in dest:
                                dest_obj = dest['/D']
                                if isinstance(dest_obj, list) and len(dest_obj) > 0:
                                    page_num = dest_obj[0] + 1 if isinstance(dest_obj[0], int) else 1
                    
                    if title:
                        result.append({
                            'title': title,
                            'page': page_num,
                            'level': level
                        })
                except:
                    pass
                
                if i + 1 < len(items) and isinstance(items[i + 1], list):
                    self._parse_outline_recursive(items[i + 1], result, level + 1)
                    i += 1
            
            i += 1
    
    def extract_text_from_page(self, page_num: int) -> str:
        if page_num < 0 or page_num >= len(self.pdf_plumber.pages):
            return ""
        
        page = self.pdf_plumber.pages[page_num]
        text = page.extract_text() or ""
        return text.strip()
    
    def is_toc_line(self, line: str) -> bool:
        """判断是否是目录行"""
        line = line.strip()
        if not line:
            return False
        
        toc_patterns = [
            r'^.{2,30}\s*…+\s*\d+$',
            r'^.{2,30}\s+\d+$',
            r'^\d+[\.\、]\s*.{2,30}\s+\d+$',
        ]
        
        for pattern in toc_patterns:
            if re.match(pattern, line):
                return True
        return False
    
    def detect_toc_pages(self) -> Tuple[int, int]:
        """检测目录页范围"""
        toc_start = -1
        toc_end = -1
        
        for i in range(min(15, len(self.pdf_plumber.pages))):
            text = self.extract_text_from_page(i)
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            
            if not lines:
                continue
            
            toc_lines = sum(1 for l in lines if self.is_toc_line(l))
            
            if toc_lines > len(lines) * 0.4:
                if toc_start == -1:
                    toc_start = i + 1
                toc_end = i + 1
            elif toc_start != -1:
                break
        
        return toc_start, toc_end
    
    def extract_all_text(self, skip_pages: set = None) -> List[Tuple[int, str]]:
        """提取所有文本，跳过指定页面"""
        texts = []
        skip_pages = skip_pages or set()
        
        for i, page in enumerate(tqdm(self.pdf_plumber.pages, desc="提取PDF文本")):
            if (i + 1) in skip_pages:
                continue
            
            text = page.extract_text() or ""
            if text.strip():
                texts.append((i + 1, text.strip()))
        
        return texts
    
    def is_chapter_title(self, line: str) -> Tuple[bool, int, str]:
        """判断是否是章节标题，返回 (是否标题, 层级, 标题)"""
        line = line.strip()
        if not line or len(line) > 60:
            return False, 0, ""
        
        if self.is_toc_line(line):
            return False, 0, ""
        
        patterns = [
            (r'^第([一二三四五六七八九十百千万零\d]+)[章节篇部]\s*(.*)$', 1),
            (r'^(\d+)[\.、\s]+(.{2,40})$', 1),
            (r'^([一二三四五六七八九十]+)[、\.]\s*(.{2,40})$', 1),
            (r'^(\d+\.\d+)\s+(.{2,40})$', 2),
            (r'^(\d+\.\d+\.\d+)\s+(.{2,40})$', 3),
            (r'^Chapter\s*(\d+)[：:\s]*(.*)$', 1),
        ]
        
        for pattern, level in patterns:
            match = re.match(pattern, line, re.IGNORECASE)
            if match:
                title = match.group(2) if match.lastindex and match.lastindex >= 2 else line
                title = re.sub(r'\s+', ' ', title).strip()
                if len(title) >= 2:
                    return True, level, title
        
        return False, 0, ""
    
    def split_into_sections(self, texts: List[Tuple[int, str]]) -> List[Section]:
        """将文本分割成章节"""
        sections = []
        current_section = None
        current_content = []
        
        for page_num, text in texts:
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                is_title, level, title = self.is_chapter_title(line)
                
                if is_title and title:
                    if current_section and current_content:
                        current_section.content = '\n'.join(current_content)
                        if len(current_section.content) >= 50:
                            sections.append(current_section)
                    
                    current_section = Section(
                        id=str(uuid.uuid4()),
                        title=title,
                        level=level,
                        content="",
                        page_start=page_num
                    )
                    current_content = []
                else:
                    if current_section:
                        current_content.append(line)
                    elif not current_section:
                        current_section = Section(
                            id=str(uuid.uuid4()),
                            title="前言",
                            level=0,
                            content="",
                            page_start=page_num
                        )
                        current_content.append(line)
        
        if current_section and current_content:
            current_section.content = '\n'.join(current_content)
            if len(current_section.content) >= 50:
                sections.append(current_section)
        
        return sections
    
    def build_section_hierarchy(self, sections: List[Section]) -> List[Section]:
        """构建章节层级结构"""
        if not sections:
            return sections
        
        root_sections = []
        section_stack = []
        
        for section in sections:
            while section_stack and section_stack[-1].level >= section.level:
                section_stack.pop()
            
            if section_stack:
                section.parent_id = section_stack[-1].id
                section_stack[-1].children.append(section)
            else:
                root_sections.append(section)
            
            section_stack.append(section)
        
        return root_sections
    
    def parse_by_outline(self) -> Optional[PDFDocument]:
        """使用PDF大纲结构解析"""
        outlines = self.get_outlines()
        
        if not outlines or len(outlines) < 3:
            return None
        
        print(f"  使用PDF大纲解析，共 {len(outlines)} 个书签")
        
        sections = []
        for i, item in enumerate(outlines):
            page_num = item.get('page', 1)
            next_page = outlines[i + 1].get('page', len(self.pdf_plumber.pages) + 1) if i + 1 < len(outlines) else len(self.pdf_plumber.pages) + 1
            
            content_parts = []
            for p in range(max(0, page_num - 1), min(next_page - 1, len(self.pdf_plumber.pages))):
                text = self.extract_text_from_page(p)
                if text:
                    content_parts.append(text)
            
            content = '\n'.join(content_parts)
            
            if len(content) >= 50:
                section = Section(
                    id=str(uuid.uuid4()),
                    title=item.get('title', ''),
                    level=item.get('level', 0),
                    content=content,
                    page_start=page_num,
                    page_end=next_page - 1
                )
                sections.append(section)
        
        if not sections:
            return None
        
        hierarchical_sections = self.build_section_hierarchy(sections)
        
        metadata = self.get_metadata()
        title = metadata.get("title", "") or self.file_path.stem
        
        return PDFDocument(
            file_path=str(self.file_path),
            title=title,
            total_pages=len(self.pdf_plumber.pages),
            sections=hierarchical_sections,
            metadata=metadata
        )
    
    def parse(self) -> PDFDocument:
        """解析PDF文档"""
        metadata = self.get_metadata()
        
        doc = self.parse_by_outline()
        if doc:
            return doc
        
        toc_start, toc_end = self.detect_toc_pages()
        skip_pages = set()
        if toc_start > 0:
            print(f"  检测到目录页: {toc_start}-{toc_end}页，已跳过")
            skip_pages = set(range(toc_start, toc_end + 1))
        
        texts = self.extract_all_text(skip_pages)
        sections = self.split_into_sections(texts)
        hierarchical_sections = self.build_section_hierarchy(sections)
        
        title = metadata.get("title", "") or self.file_path.stem
        
        document = PDFDocument(
            file_path=str(self.file_path),
            title=title,
            total_pages=len(self.pdf_plumber.pages),
            sections=hierarchical_sections,
            metadata=metadata
        )
        
        return document
    
    def get_all_sections_flat(self, sections: List[Section]) -> List[Section]:
        result = []
        for section in sections:
            result.append(section)
            if section.children:
                result.extend(self.get_all_sections_flat(section.children))
        return result
