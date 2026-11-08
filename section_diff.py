"""作者：Wangyuan；日期：2026-10-29。职责：通用PDF章节/字段差异和显式文件组对比。"""
from __future__ import annotations
from dataclasses import dataclass, field, asdict
from pathlib import Path
import difflib
import re
from pdf_text import extract_pdf


@dataclass
class DiffRules:
    """章节、字段与噪声由调用方配置；无原业务资料或噪声规则。"""
    section_pattern: str = r'^##\s+(.+)$'
    label_pattern: str | None = None
    data_pattern: str = r'^\d+\s'
    noise_patterns: list[str] = field(default_factory=list)
    lookback: int = 20

    def __post_init__(self) -> None:
        if type(self.lookback) is not int or not 1 <= self.lookback <= 100: raise ValueError('回溯预算无效')
        if not isinstance(self.noise_patterns,list) or len(self.noise_patterns)>30: raise ValueError('噪声规则超出预算')
        for pattern in [self.section_pattern,self.data_pattern,*self.noise_patterns]+([self.label_pattern] if self.label_pattern else []):
            if not isinstance(pattern,str) or len(pattern)>200: raise ValueError('表达式无效或过长')
            try: re.compile(pattern)
            except re.error as exc: raise ValueError("表达式无效") from exc
        if re.compile(self.section_pattern).groups<1: raise ValueError('章节表达式须有标题捕获组')

    def is_noise(self, line: str) -> bool:
        """只有显式噪声表达式命中才过滤。"""
        return any(re.search(pattern,line) for pattern in self.noise_patterns)

    def section(self, line: str) -> str:
        """识别用户表达式中的标题；没有命中则返回空。"""
        match=re.match(self.section_pattern,line)
        return (match.group(1) or '').strip() if match else ''

    def is_label(self, line: str) -> bool:
        """保留原中文短标签启发式，也接受显式标签表达式。"""
        if self.label_pattern: return re.fullmatch(self.label_pattern,line) is not None
        count=len(re.findall(r'[\u4e00-\u9fff]',line))
        return count>=2 and count/max(len(line),1)>=0.4 and len(line)<=30 and not re.search(r'\d{3,}',line) and not re.match(r'^[A-Z]\s',line)


def nearest_field_label(lines: list[str], index: int, rules: DiffRules) -> str:
    """回溯最近合格标签，空文档不越界。"""
    for cursor in range(min(index,len(lines)-1),max(-1,index-rules.lookback),-1):
        line=lines[cursor]
        if rules.section(line):break
        if rules.is_noise(line) or re.match(rules.data_pattern,line): continue
        if rules.is_label(line): return line
    return ''


@dataclass
class FileChange:
    """保留变更章节、字段、原文差异块及质量边界。"""
    file_name: str
    sections: dict[str,list[str]] = field(default_factory=dict)
    hunks: list[dict] = field(default_factory=list)

    def add(self, section: str, label: str) -> None:
        """稳定保序去重，不把集合迭代顺序冒充文档顺序。"""
        labels=self.sections.setdefault(section or '未识别章节',[])
        if label and label not in labels: labels.append(label)


def diff_lines(old: list[str], new: list[str], rules: DiffRules | None = None, name: str = 'explicit-document') -> FileChange:
    """按原SequenceMatcher算法提取差异，纯删除采用旧文档章节。"""
    rules=rules or DiffRules()
    for lines in (old,new):
        if not isinstance(lines,list) or len(lines)>5000 or any(not isinstance(line,str) or len(line)>2000 for line in lines) or sum(map(len,lines))>200000:
            raise ValueError('文本行数或字符超过预算')
    old=[line.strip() for line in old if line.strip() and not rules.is_noise(line)]
    new=[line.strip() for line in new if line.strip() and not rules.is_noise(line)]
    def sections(lines: list[str]) -> list[str]:
        current=''; result=[]
        for line in lines:
            current=rules.section(line) or current;result.append(current)
        return result
    before,after=sections(old),sections(new); result=FileChange(name)
    for tag,i1,i2,j1,j2 in difflib.SequenceMatcher(a=old,b=new,autojunk=False).get_opcodes():
        if tag=='equal':continue
        source,index,section_map=(old,i1,before) if tag=='delete' else (new,j1,after)
        section=section_map[index] if index<len(section_map) else ''
        header_change=index<len(source) and bool(rules.section(source[index]))
        label='' if header_change else nearest_field_label(source,index-1,rules)
        old_section=before[min(i1,len(before)-1)] if before else ''
        if not label and not header_change and tag!='delete' and old_section==section:
            label=nearest_field_label(old,i1-1,rules)
        result.add(section,label)
        affected=[]
        # 单个opcode可跨多个章节；新旧两侧都登记，避免整段删除章节被首行遮蔽。
        for lines,start,end,positions in ((old,i1,i2,before),(new,j1,j2,after)):
            previous=None
            for cursor in range(start,end):
                title=positions[cursor]
                if title!=previous:
                    field=nearest_field_label(lines,cursor-1,rules)
                    if rules.section(lines[cursor]):field=''
                    result.add(title,field)
                    if title not in affected:affected.append(title)
                    previous=title
        result.hunks.append({'kind':tag,'section':section,'affected_sections':affected,'field':label,'old_lines':old[i1:i2],'new_lines':new[j1:j2]})
    return result


def extract_lines(path: Path) -> list[str]:
    """调用有超时的文字层提取；空文字页不能假装已完成全文比较。"""
    result=extract_pdf(path)
    if result['empty_text_pages'] or not result['pages']: raise ValueError('存在空文字页或扫描页，需校对文字后比较')
    return [line.strip() for page in result['pages'] for line in page['text'].splitlines() if line.strip()]


def diff_files(old_path: str | Path, new_path: str | Path, rules: DiffRules | None = None) -> FileChange:
    """比较两个显式PDF，仅覆盖可提取文字层，不比较图形。"""
    old_path,new_path=Path(old_path),Path(new_path)
    return diff_lines(extract_lines(old_path),extract_lines(new_path),rules,new_path.name)


def compare_group(label: str, old_dir: str | Path, new_dir: str | Path, rules: dict | None = None) -> dict:
    """比较两个显式目录内同名PDF；错误单列为不完整，不伪作无差异。"""
    config=DiffRules(**(rules or {})); old_dir,new_dir=Path(old_dir),Path(new_dir)
    if not old_dir.is_dir() or not new_dir.is_dir():raise ValueError('新旧输入目录必须存在')
    def files(directory: Path) -> dict[str,Path]:
        result={}
        for path in directory.iterdir():
            if path.suffix.lower()!='.pdf':continue
            if path.is_symlink():raise ValueError('PDF目录中含符号链接，请提供实际文件')
            if path.is_file():result[path.name]=path
        if len(result)>100:raise ValueError('目录超过100个PDF预算')
        return result
    old,new=files(old_dir),files(new_dir)
    result={'label':label,'only_in_new':sorted(new.keys()-old.keys()),'only_in_old':sorted(old.keys()-new.keys()),'file_changes':[],'errors':[],'coverage':'text_layer_only'}
    for name in sorted(old.keys()&new.keys()):
        try:
            change=diff_files(old[name],new[name],config)
            if change.hunks:result['file_changes'].append(asdict(change))
        except (ValueError,OSError) as exc:
            result['errors'].append({'file':name,'error':str(exc)})
    result['state']='incomplete' if result['errors'] else 'review_required'
    result['has_changes']=bool(result['file_changes'] or result['only_in_new'] or result['only_in_old'])
    return result
