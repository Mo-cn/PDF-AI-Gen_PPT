"""
PDF-AI-GEN_PPT 主程序入口
"""
import click
import json
from pathlib import Path
from datetime import datetime
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from .pdf_parser import PDFParser
from .question_generator import QuestionGenerator
from .ppt_generator import PPTGenerator
from .output_manager import OutputManager
from .config import settings, AIProvider
from .models import PDFDocument, QuestionSet, PPTDocument, Question, QuestionType, Difficulty

console = Console()


@click.group()
@click.version_option(version="1.0.0", prog_name="PDF-AI-GEN_PPT")
def cli():
    """PDF-AI-GEN_PPT: 基于AI接口的PDF内容处理与PPT自动生成系统"""
    pass


def save_questions_incremental(question_sets: list, output_dir: str, filename: str = None):
    """增量保存试题"""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    filename = filename or f"questions_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    json_path = output_path / f"{filename}.json"
    
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
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    return str(json_path)


@cli.command()
@click.argument('pdf_path', type=click.Path(exists=True))
@click.option('--output', '-o', default='output', help='输出目录')
@click.option('--questions', '-q', default=25, help='每节生成的题目数量')
@click.option('--no-ppt', is_flag=True, help='不生成PPT')
@click.option('--no-questions', is_flag=True, help='不生成试题')
@click.option('--combined-ppt', is_flag=True, help='生成综合PPT而非分章节PPT')
@click.option('--use-ai/--no-ai', default=True, help='是否使用AI生成PPT内容')
@click.option('--section', '-s', default=None, help='指定章节编号(如: 1,3,5-8)')
def process(pdf_path, output, questions, no_ppt, no_questions, combined_ppt, use_ai, section):
    """处理PDF文件，生成试题和PPT"""
    
    console.print(Panel.fit(
        "[bold blue]PDF-AI-GEN_PPT[/bold blue]\n"
        "基于AI接口的PDF内容处理与PPT自动生成系统",
        border_style="blue"
    ))
    
    if not settings.AI_API_KEY:
        console.print("[red]错误: 未设置AI_API_KEY环境变量[/red]")
        return
    
    output_dir = Path(output)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print("\n[1/3] 解析PDF文件...")
    with PDFParser(pdf_path) as parser:
        document = parser.parse()
    print(f"  ✓ 完成: {document.total_pages}页\n")
    
    output_manager = OutputManager(output)
    structure_path = output_manager.save_document_structure(document)
    
    question_sets = []
    if not no_questions:
        print("[2/3] 生成试题...")
        generator = QuestionGenerator()
        
        from .question_generator import _interrupted
        
        flat_sections = generator._flatten_sections(document.sections)
        valid_sections = [s for s in flat_sections if len(s.content.strip()) >= 100]
        
        if section:
            valid_sections = filter_sections_by_range(valid_sections, section)
            print(f"  指定处理 {len(valid_sections)} 个章节\n")
        else:
            print(f"  共 {len(valid_sections)} 个章节需要处理")
            print("  每完成一个章节会自动保存\n")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        for i, sec in enumerate(valid_sections, 1):
            if _interrupted:
                break
            
            print(f"  [{i}/{len(valid_sections)}] {sec.title[:30]}...")
            try:
                question_set = generator.generate_for_section(
                    section=sec,
                    num_questions=questions
                )
                question_sets.append(question_set)
                print(f"    ✓ 生成 {question_set.total_count} 题")
                
                save_questions_incremental(question_sets, output, f"questions_{timestamp}")
                
            except KeyboardInterrupt:
                print("\n  中断，已保存当前进度")
                break
            except Exception as e:
                print(f"    ✗ 失败: {str(e)[:80]}")
                continue
        
        if question_sets:
            print(f"\n  试题生成完成: {len(question_sets)} 章节, {sum(qs.total_count for qs in question_sets)} 题")
            json_path = save_questions_incremental(question_sets, output, f"questions_{timestamp}")
            excel_paths = output_manager.save_questions_to_excel(question_sets, f"questions_{timestamp}", separate_answer=True)
            print(f"  JSON: {json_path}")
            print(f"  题目: {excel_paths.get('questions')}")
            print(f"  答案: {excel_paths.get('answers')}\n")
    
    ppt_documents = []
    if not no_ppt:
        print("[3/3] 生成PPT...")
        ppt_generator = PPTGenerator()
        
        if combined_ppt:
            combined_path = output_dir / f"{document.title}_综合.pptx"
            ppt_path = ppt_generator.generate_combined_ppt(
                document.sections,
                str(combined_path),
                use_ai=use_ai
            )
            print(f"  ✓ PPT已保存: {ppt_path}")
        else:
            ppt_output_dir = output_dir / "ppt"
            ppt_documents = ppt_generator.generate_for_all_sections(
                document.sections,
                str(ppt_output_dir),
                use_ai=use_ai
            )
            print(f"  ✓ 生成 {len(ppt_documents)} 个PPT文件\n")
    
    report_path = output_manager.generate_report(document, question_sets, ppt_documents)
    
    console.print("\n[green]✓ 处理完成！[/green]")
    console.print(f"报告: {report_path}")


def filter_sections_by_range(sections, range_str):
    """根据范围字符串过滤章节"""
    selected = []
    ranges = range_str.split(',')
    
    for r in ranges:
        r = r.strip()
        if '-' in r:
            start, end = map(int, r.split('-'))
            for i in range(start - 1, min(end, len(sections))):
                selected.append(sections[i])
        else:
            idx = int(r) - 1
            if 0 <= idx < len(sections):
                selected.append(sections[idx])
    
    return selected


@cli.command()
@click.argument('pdf_path', type=click.Path(exists=True))
@click.option('--output', '-o', default='output', help='输出目录')
def parse(pdf_path, output):
    """仅解析PDF文件结构"""
    
    print("解析PDF文件...\n")
    
    with PDFParser(pdf_path) as parser:
        document = parser.parse()
    
    output_manager = OutputManager(output)
    structure_path = output_manager.save_document_structure(document)
    
    print(f"✓ 解析完成")
    print(f"文档标题: {document.title}")
    print(f"总页数: {document.total_pages}")
    print(f"结构文件: {structure_path}\n")
    
    flat_sections = []
    def collect_sections(sections, level=0):
        for s in sections:
            flat_sections.append((len(flat_sections) + 1, s.title, level))
            if s.children:
                collect_sections(s.children, level + 1)
    
    collect_sections(document.sections)
    
    print("章节列表:")
    for idx, title, level in flat_sections:
        indent = "  " * level
        print(f"{indent}[{idx}] {title}")


@cli.command()
@click.argument('pdf_path', type=click.Path(exists=True))
@click.option('--output', '-o', default='output', help='输出目录')
@click.option('--questions', '-q', default=25, help='每节生成的题目数量')
@click.option('--section', '-s', default=None, help='指定章节编号(如: 1,3,5-8)')
@click.option('--retry', '-r', default=0, help='失败重试次数')
def questions(pdf_path, output, questions, section, retry):
    """仅生成试题
    
    示例:
      python run.py questions one.pdf                    # 处理所有章节
      python run.py questions one.pdf -s 1,3,5           # 处理第1,3,5章
      python run.py questions one.pdf -s 5-10            # 处理第5到10章
      python run.py questions one.pdf -s 1,3,5-8 -r 2    # 指定章节并重试2次
    """
    
    if not settings.AI_API_KEY:
        console.print("[red]错误: 未设置AI_API_KEY环境变量[/red]")
        return
    
    print("解析PDF并生成试题...\n")
    
    with PDFParser(pdf_path) as parser:
        document = parser.parse()
    
    generator = QuestionGenerator()
    
    from .question_generator import _interrupted
    
    flat_sections = generator._flatten_sections(document.sections)
    valid_sections = [s for s in flat_sections if len(s.content.strip()) >= 100]
    
    if section:
        valid_sections = filter_sections_by_range(valid_sections, section)
        print(f"指定处理 {len(valid_sections)} 个章节\n")
    else:
        print(f"共 {len(valid_sections)} 个章节需要处理\n")
    
    question_sets = []
    failed_sections = []
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_manager = OutputManager(output)
    
    for i, sec in enumerate(valid_sections, 1):
        if _interrupted:
            break
        
        print(f"[{i}/{len(valid_sections)}] {sec.title[:30]}...")
        
        success = False
        for attempt in range(retry + 1):
            if attempt > 0:
                print(f"    重试 {attempt}/{retry}...")
            
            try:
                question_set = generator.generate_for_section(
                    section=sec,
                    num_questions=questions
                )
                
                if question_set.total_count > 0:
                    question_sets.append(question_set)
                    print(f"  ✓ 生成 {question_set.total_count} 题")
                    save_questions_incremental(question_sets, output, f"questions_{timestamp}")
                    success = True
                    break
                else:
                    print(f"  ✗ 生成0题")
            except Exception as e:
                print(f"  ✗ 失败: {str(e)[:60]}")
        
        if not success:
            failed_sections.append((i, sec.title))
    
    if question_sets:
        print(f"\n✓ 完成: {len(question_sets)} 章节, {sum(qs.total_count for qs in question_sets)} 题")
        json_path = save_questions_incremental(question_sets, output, f"questions_{timestamp}")
        excel_paths = output_manager.save_questions_to_excel(question_sets, f"questions_{timestamp}", separate_answer=True)
        print(f"JSON: {json_path}")
        print(f"题目: {excel_paths.get('questions')}")
        print(f"答案: {excel_paths.get('answers')}")
    
    if failed_sections:
        print(f"\n失败章节 ({len(failed_sections)}):")
        for idx, title in failed_sections:
            print(f"  [{idx}] {title}")


@cli.command()
@click.argument('pdf_path', type=click.Path(exists=True))
@click.option('--output', '-o', default='output', help='输出目录')
@click.option('--combined', is_flag=True, help='生成综合PPT')
@click.option('--use-ai/--no-ai', default=True, help='是否使用AI生成PPT内容')
def ppt(pdf_path, output, combined, use_ai):
    """仅生成PPT"""
    
    if not settings.AI_API_KEY:
        console.print("[red]错误: 未设置AI_API_KEY环境变量[/red]")
        return
    
    print("解析PDF并生成PPT...\n")
    
    with PDFParser(pdf_path) as parser:
        document = parser.parse()
    
    output_dir = Path(output)
    ppt_generator = PPTGenerator()
    
    if combined:
        combined_path = output_dir / f"{document.title}_综合.pptx"
        ppt_path = ppt_generator.generate_combined_ppt(
            document.sections,
            str(combined_path),
            use_ai=use_ai
        )
        print(f"\n✓ PPT已保存: {ppt_path}")
    else:
        ppt_output_dir = output_dir / "ppt"
        ppt_documents = ppt_generator.generate_for_all_sections(
            document.sections,
            str(ppt_output_dir),
            use_ai=use_ai
        )
        print(f"\n✓ 生成 {len(ppt_documents)} 个PPT文件")


@cli.command()
@click.argument('json_path', type=click.Path(exists=True))
@click.option('--output', '-o', default='output', help='输出目录')
def export(json_path, output):
    """从JSON文件导出Excel格式"""
    
    output_manager = OutputManager(output)
    question_sets = output_manager.load_questions_from_json(json_path)
    
    if not question_sets:
        print("未找到试题数据")
        return
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_paths = output_manager.save_questions_to_excel(question_sets, f"export_{timestamp}", separate_answer=True)
    
    print(f"导出完成:")
    print(f"题目: {excel_paths.get('questions')}")
    print(f"答案: {excel_paths.get('answers')}")


@cli.command()
def config():
    """显示当前配置"""
    
    table = Table(title="当前配置")
    table.add_column("配置项", style="cyan")
    table.add_column("值", style="green")
    
    table.add_row("AI提供商", settings.AI_PROVIDER.value)
    table.add_row("AI模型", settings.AI_MODEL)
    table.add_row("API密钥", "***已设置***" if settings.AI_API_KEY else "未设置")
    table.add_row("每节题目数", str(settings.QUESTIONS_PER_SECTION))
    table.add_row("输出目录", settings.OUTPUT_DIR)
    
    console.print(table)


def main():
    cli()


if __name__ == '__main__':
    main()
