#!/usr/bin/env python
# -*- coding: utf-8 -*-

import builtins
import sys, os, json, traceback

# ========================================================
# 核心修正：劫持全局 print 并强制刷新 stdout，同时记录到全局日志
# 解决 Sigil 插件面板缓冲问题，并允许将控制台内容写入 error_log
# ========================================================
console_logs = []

def _flush_print(*args, **kwargs):
    sep = kwargs.get('sep', ' ')
    msg = sep.join(map(str, args))
    console_logs.append(msg)
    builtins.print(*args, **kwargs)
    sys.stdout.flush()

print = _flush_print

try:
    import regex as re
except ImportError:
    import re

# 引入 Sigil 环境内置的 BeautifulSoup4 库，大幅简化 XML/HTML 的解析
try:
    from bs4 import BeautifulSoup
except ImportError:
    pass

import lxml.html
from lxml import etree

# ========================================================
# 修复跨设备/不同Sigil版本导致的 Qt platform plugin 冲突问题
# ========================================================
for _env_key in ['QT_PLUGIN_PATH', 'QT_QPA_PLATFORM_PLUGIN_PATH']:
    if _env_key in os.environ:
        del os.environ[_env_key]

# 动态加载 Qt 库
try:
    from PySide6 import QtWidgets, QtCore, QtGui
except ImportError:
    try:
        from PyQt6 import QtWidgets, QtCore, QtGui
    except ImportError:
        from PyQt5 import QtWidgets, QtCore, QtGui

class BBCodeConverter:
    def __init__(self, bk):
        self.bk = bk
        self.footnote_map = {}
        self.nav_titles = []
        self.nav_uid = ""
        self.nav_filename = ""
        self.skip_files = set()
        self.toc_files = {}
        self.spine_map = {}

        # 高频复杂正则表达式预编译缓存
        self.PAT_A_TAG = re.compile(r'<a[^>]*href=["\']#([^"\']+)["\'][^>]*>(.*?)</a>', flags=re.I|re.S)
        self.PAT_RUBY = re.compile(r'<ruby[^>]*>(.*?)</ruby>', flags=re.I|re.S)
        self.PAT_RUBY_RT = re.compile(r'<rt[^>]*>(.*?)</rt>', flags=re.I|re.S)
        self.PAT_IMG = re.compile(r'<(?:img|image)[^>]+(?:src|href)=["\']([^"\']+)["\'][^>]*>', flags=re.I)

    def add_nav_title(self, title, href_full):
        if not title or not href_full: return
        filename = os.path.basename(href_full.split('#')[0])
        title = title.strip(' \t\n\r\u3000')
        
        if title in ["書封", "书封", "封面", "作者頁", "作者页", "書名頁", "书名页", "內彩", "封底", "彩頁", "彩页"]:
            print(f"  -> [识别-忽略] 发现封面标题: '{title}'，不将其纳入匹配库（但保留页面正文读取）")
            return
            
        target_idx = self.spine_map.get(filename, 9999)
        
        if title in ["版權頁", "版权页"]:
            self.skip_files.add(filename)
            print(f"  -> [识别-剥离] 锁定特权页面并打上跳过标签: {filename} (类型: '{title}')")
        elif title in ["目录", "目錄", "目 录", "目 錄", "目　录", "目　錄"]:
            self.toc_files[filename] = title
            print(f"  -> [识别-目录] 锁定目录页面: {filename} (将为其整页添加首尾居中标签)")
        else:
            self.nav_titles.append({
                'filename': filename, 
                'title': title, 
                'matched': False,
                'target_idx': target_idx
            })
            print(f"  -> [识别-导航] 规划定位点: {filename} | 预期标题: '{title}' | 所属OPF索引位: {target_idx}")

    def pre_scan(self):
        print("\n" + "="*50 + "\n【第一步：解析 OPF 导航与构建全局防逆流书脊地图】\n" + "="*50)
        try:
            manifest_map = {}
            for m_uid, m_href, m_mime in self.bk.manifest_iter():
                manifest_map[m_uid] = m_href

            for idx, spine_info in enumerate(self.bk.spine_iter()):
                uid = spine_info[0]
                href = manifest_map.get(uid, "")
                if href:
                    self.spine_map[os.path.basename(href.split('#')[0])] = idx

            opf_html = self.bk.get_opf()
            soup = BeautifulSoup(opf_html, 'html.parser')
            
            nav_item = soup.find('item', properties=lambda x: x and 'nav' in x.lower())
            m_toc = soup.find('spine')
            
            self.nav_uid = nav_item['id'] if nav_item else (m_toc.get('toc') if m_toc and m_toc.get('toc') else "toc")
            
            for uid, href, mime in self.bk.manifest_iter():
                if uid == self.nav_uid:
                    self.nav_filename = os.path.basename(href or "")
                    print(f"[*] 成功锁定官方导航文件: {self.nav_filename} (UID: {self.nav_uid})")
                    break
        except Exception as e:
            print(f"[!] OPF 解析发生异常: {e}")

        if self.nav_uid:
            try:
                nav_html = self.bk.readfile(self.nav_uid)
                print("[*] 开始使用 BeautifulSoup 从导航中提取预设标题...")
                nav_soup = BeautifulSoup(nav_html, 'html.parser')
                
                if self.nav_filename.lower().endswith('.ncx'):
                    for navpoint in nav_soup.find_all('navpoint'):
                        text_tag = navpoint.find('text')
                        content_tag = navpoint.find('content')
                        if text_tag and content_tag and content_tag.has_attr('src'):
                            self.add_nav_title(text_tag.get_text(strip=True), content_tag['src'])
                else:
                    toc_nav = nav_soup.find('nav', attrs={'epub:type': 'toc'})
                    if not toc_nav:
                        toc_nav = nav_soup
                    
                    for a_tag in toc_nav.find_all('a', href=True):
                        self.add_nav_title(a_tag.get_text(strip=True), a_tag['href'])
            except Exception as e:
                print(f"[!] 导航解析发生异常: {e}")

        try:
            print("[*] 开始使用 BeautifulSoup 提取全书备用脚注...")
            for text_info in self.bk.text_iter():
                uid, href = text_info[0], text_info[1] if len(text_info)>1 else ""
                filename = os.path.basename(href or "")
                if uid == self.nav_uid or filename == self.nav_filename or filename in self.skip_files or filename in self.toc_files: 
                    continue
                
                content = self.bk.readfile(uid)
                soup = BeautifulSoup(content, 'html.parser')
                
                for aside in soup.find_all('aside'):
                    for li in aside.find_all('li', id=True):
                        txt = li.get_text(strip=True).strip(' \t\n\r\u3000')
                        if txt: self.footnote_map[li['id']] = txt
                        
                for fn in soup.find_all(['p', 'div', 'span', 'li'], class_=re.compile(r'footnote', re.I), id=True):
                    txt = fn.get_text(strip=True).strip(' \t\n\r\u3000')
                    if txt: self.footnote_map[fn['id']] = txt
                    
            print(f"[*] 共提取到 {len(self.footnote_map)} 个备用脚注。")
        except Exception as e:
            print(f"[!] 备用脚注提取发生异常: {e}")

    def clean_and_convert(self, html, href, img_map, deleted_imgs, current_idx, inject_title=False):
        href = href or ""
        filename = os.path.basename(href.split('#')[0])
        
        expected_titles = [nt for nt in self.nav_titles if nt['filename'] == filename and not nt['matched'] and current_idx >= nt['target_idx']]

        print(f"\n--- 【文件处理开始】: {filename} (当页绝对索引: {current_idx}) ---")
        if expected_titles:
            print(f"  [监测-标题池] 本页准许并激活匹配 {len(expected_titles)} 个标题。")

        if re.search(r'<table\b', html, flags=re.I):
            print(f"  [警告] 侦测到表格标签 <table>，当前版本不支持表格转换，请注意输出排版。")

        html = re.sub(r'>\s*\n\s*<', '><', html)

        print(f"  [处理-注释] 开始扫描当页脚注定义...")
        local_footnotes = {}
        del_blocks = []
        
        a_tags = self.PAT_A_TAG.finditer(html)
        
        for m in a_tags:
            link_pos = m.start()
            target_id = m.group(1)
            a_text_raw = re.sub(r'<[^>]+>', '', m.group(2)).strip(' \t\n\r\u3000[]()【】')
            
            if '注' in a_text_raw or '註' in a_text_raw or '*' in a_text_raw or a_text_raw.isdigit():
                target_str1 = f'id="{target_id}"'
                target_str2 = f"id='{target_id}'"
                
                idx = html.find(target_str1)
                if idx == -1: idx = html.find(target_str2)
                
                if idx != -1 and idx > link_pos:
                    block_start = -1
                    tag_name = ""
                    for tag in ["p", "div", "li", "aside", "section"]:
                        temp_start = html.rfind(f"<{tag}", 0, idx)
                        if temp_start > block_start:
                            block_start = temp_start
                            tag_name = tag
                            
                    if block_start != -1:
                        block_end = html.find(f"</{tag_name}>", idx)
                        if block_end != -1:
                            block_end += len(f"</{tag_name}>")
                            block_html = html[block_start:block_end]
                            
                            txt = re.sub(r'<[^>]+>', '', block_html).strip(' \t\n\r\u3000')
                            txt = re.sub(r'\s+', ' ', txt)
                            txt = re.sub(r'^[\^↵↑*※]\s*', '', txt)
                            txt = re.sub(r'^返回正文\s*', '', txt)
                            txt = re.sub(r'^' + re.escape(a_text_raw) + r'[:：\s]*', '', txt)
                            
                            local_footnotes[target_id] = {'text': txt}
                            if block_html not in del_blocks:
                                del_blocks.append(block_html)

        for block in del_blocks:
            if block in html:
                html = html.replace(block, f'[SYS_DEL_FOOTNOTE]{block}[/SYS_DEL_FOOTNOTE]')
        if del_blocks:
            print(f"  [处理-注释销毁准备] 已为 {len(del_blocks)} 个底部注释区块添加删除标识符。")

        def footnote_link_replacer(m):
            target_id = m.group(1)
            original_a_tag = m.group(0)
            if target_id in local_footnotes:
                print(f"    [匹配-注释] 成功拉取注释: #{target_id}")
                return f"（{local_footnotes[target_id]['text']}）"
            elif target_id in self.footnote_map:
                print(f"    [匹配-注释] 从全局备用拉取: #{target_id}")
                return f"（{self.footnote_map[target_id]}）"
            else:
                return original_a_tag
                
        html, fn_count = self.PAT_A_TAG.subn(footnote_link_replacer, html)
        if fn_count > 0:
            print(f"  [处理-注释] 发现并内联了 {fn_count} 处脚注引用。")

        html, del_count = re.subn(r'\[SYS_DEL_FOOTNOTE\].*?\[/SYS_DEL_FOOTNOTE\]', '', html, flags=re.I|re.S)
        if del_count > 0:
            print(f"  [处理-注释销毁执行] 成功清除了 {del_count} 个被标记的底部注释区块。")

        while True:
            new_html = re.sub(r'<([a-z0-9]+)[^>]*>\s*</\1>', '', html, flags=re.I)
            if new_html == html: break
            html = new_html

        html = re.sub(r'</rt>\s+</ruby>', '</rt></ruby>', html, flags=re.I)
        def ruby_repl(m):
            inner = m.group(1)
            rts = "".join([re.sub(r'<[^>]+>', '', c) for c in self.PAT_RUBY_RT.findall(inner)])
            base = re.sub(r'<(?:rp|rt)[^>]*>.*?</(?:rp|rt)>', '', inner, flags=re.I|re.S)
            base = re.sub(r'<[^>]+>', '', base).replace('\n', '')
            return f"[ruby={rts}]{base}[/ruby]"
            
        html, r_count = self.PAT_RUBY.subn(ruby_repl, html)
        if r_count > 0: print(f"  [处理-Ruby] 成功转换了 {r_count} 处 Ruby 注音。")

        html, b_count = re.subn(r'<p[^>]*>(.*?)</p>', lambda m: "<p>" + re.sub(r'<br\s*/?>', '[SYS_BR_SPACE]', m.group(1), flags=re.I) + "</p>", html, flags=re.S|re.I)
        
        def img_repl(m):
            src = os.path.basename(m.group(1))
            return "" if src in deleted_imgs else img_map.get(src, f"__IMG_MARKER__{src}__")
            
        html, i_count = self.PAT_IMG.subn(img_repl, html)
        if i_count > 0: print(f"  [处理-图片] 成功映射了 {i_count} 张图片。")

        html = re.sub(r'<hr[^>]*>', '\n[SYS_HR_MARKER]\n', html, flags=re.I)

        # ========================================================
        # 【多阶样式统一引擎 - LXML DOM 树解析重构版】
        # 彻底抛弃低效嵌套正则，基于内存树底向上解析，性能飞跃并杜绝回溯死循环
        # ========================================================
        if html.strip():
            try:
                root = lxml.html.fromstring(f"<div id='__sys_root__'>{html}</div>")
                
                # 安全包裹函数：在不破坏节点原有层级的前提下，在元素首尾追加 BBCode 标识符
                def wrap_contents(el, start_tag, end_tag):
                    if el.text:
                        el.text = start_tag + el.text
                    else:
                        el.text = start_tag
                        
                    if len(el) > 0:
                        last_child = el[-1]
                        if last_child.tail:
                            last_child.tail += end_tag
                        else:
                            last_child.tail = end_tag
                    else:
                        el.text += end_tag

                # 核心解析：使用 reversed 实现从最内层向最外层的倒序（Bottom-Up）遍历
                for el in reversed(list(root.iter())):
                    if el.tag == 'div' and el.get('id') == '__sys_root__':
                        continue
                    if not isinstance(el.tag, str):
                        continue
                        
                    tag = el.tag.lower()
                    if tag not in ('span', 'b', 'strong', 'i', 'em', 's', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                        continue
                        
                    # 清理因源码排版产生的多余换行，防止 BBCode 标签产生跨行包裹错误
                    if el.text:
                        el.text = el.text.lstrip('\r\n')
                    if len(el) > 0:
                        last_child = el[-1]
                        if last_child.tail:
                            last_child.tail = last_child.tail.rstrip('\r\n')
                    elif el.text:
                        el.text = el.text.rstrip('\r\n')
                        
                    attrs_str = " ".join([f"{k}='{v}'" for k, v in el.attrib.items()]).lower()
                    
                    if tag in ('b', 'strong'):
                        wrap_contents(el, '[b]', '[/b]')
                    elif tag in ('i', 'em'):
                        wrap_contents(el, '[i]', '[/i]')
                    elif tag == 's':
                        wrap_contents(el, '[s]', '[/s]')
                        
                    if re.search(r'(bold|gfont|font-weight:\s*bold)', attrs_str):
                        wrap_contents(el, '[b]', '[/b]')
                    if re.search(r'(italic|font-style:\s*italic)', attrs_str):
                        wrap_contents(el, '[i]', '[/i]')
                    if re.search(r'(line-through|strike|text-decoration:\s*line-through)', attrs_str):
                        wrap_contents(el, '[s]', '[/s]')
                    if re.search(r'(align-center|text-center|text-align:\s*center)', attrs_str):
                        wrap_contents(el, '[center]', '[/center]')
                    if re.search(r'(align-left|align-start|text-left|text-start|text-align:\s*(left|start))', attrs_str):
                        wrap_contents(el, '[left]', '[/left]')
                    if re.search(r'(align-right|align-end|text-right|text-end|text-align:\s*(right|end))', attrs_str):
                        wrap_contents(el, '[right]', '[/right]')
                        
                    if re.search(r'start-6em', attrs_str):
                        inner_text = "".join(el.itertext()).strip(' \t\n\r\u3000')
                        inner_text = re.sub(r'\s+', ' ', inner_text)
                        for child in list(el):
                            el.remove(child)
                        el.text = f"[segmentation]{inner_text}[/segmentation]"

                    if tag in ('p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                        sys_tag = tag
                        
                        if tag == 'div':
                            # 检查是否包含系统级的块级标签（因为是从内向外处理，子块级已被更名为 sys_ 系列）
                            has_sys_child = any(isinstance(c.tag, str) and c.tag.startswith('sys_') for c in el.iterdescendants())
                            if not has_sys_child:
                                log_text = "".join(el.itertext()).strip(' \t\n\r\u3000')
                                log_text = re.sub(r'\[.*?\]', '', log_text)
                                if log_text:
                                    if len(log_text) > 20: log_text = log_text[:20] + '...'
                                    print(f"  [检查提醒] 发现直接包裹文本的 div 标签，已转换换行标记: '{log_text}'")
                                    sys_tag = 'div_nl'
                                
                        el.tag = f"sys_{sys_tag}"
                        el.attrib.clear()
                    else:
                        # 行内标签仅剥离外壳，将其处理完毕的 BBCode 结构无缝并入父级或前后兄弟节点的字符流中
                        el.drop_tag()
                        
                # 重新序列化提取处理完毕的内部 HTML 结构
                html = (root.text or "") + "".join(etree.tostring(child, encoding='unicode', method='html') for child in root)
            except Exception as e:
                print(f"  [致命警告] LXML 引擎解析当前页面崩溃，尝试跳过样式处理: {e}")

        html = re.sub(r'</sys_p>|</sys_h[1-6]>|</sys_div_nl>|</li>|<br\s*/?>', '\n', html, flags=re.I)
        text = re.sub(r'<[^>]+>', '', html)

        raw_lines = text.split('\n')
        marked_lines = []

        print(f"  [处理-标题比对] 开始进入核心按行扫描...")
        for line_num, line in enumerate(raw_lines, 1):
            stripped_line = line.strip(' \t\r\n\u3000')
            if not stripped_line:
                marked_lines.append(line)
                continue

            if '[segmentation]' in stripped_line or '[SYS_HR_MARKER]' in stripped_line:
                marked_lines.append(line)
                continue

            temp_line = re.sub(r'\[img\].*?\[/img\]', '', stripped_line, flags=re.I)
            temp_line = re.sub(r'__IMG_MARKER__.*?__', '', temp_line)

            pure_text = re.sub(r'\[/?[a-z=0-9]+\]', '', temp_line).strip(' \t\r\n\u3000')
            norm_line = re.sub(r'\s+', '', pure_text)
            match_found = False

            if norm_line and expected_titles:
                for nt in expected_titles:
                    if not nt['matched']:
                        exp_norm = re.sub(r'\s+', '', nt['title'])
                        
                        if len(norm_line) > 1 and (norm_line in exp_norm or exp_norm in norm_line):
                             print(f"    [分析追踪 | 行 {line_num:03d}] 探测到包含标题核心字的行。目标:'{exp_norm}' | 剥离后提取为:'{norm_line}'")
                        
                        if norm_line == exp_norm:
                            print(f"    ✅ [命中-成功锁定首出位 | 行 {line_num:03d}] 原文提取的 '{pure_text}' == 导航标题 '{nt['title']}'。已加锁。")
                            marked_lines.append(f"[SYS_TITLE]{nt['title']}[/SYS_TITLE]")
                            nt['matched'] = True
                            match_found = True
                            break

            if not match_found:
                marked_lines.append(line)

        result_lines = marked_lines
        if expected_titles:
            print(f"  [章节战报结算] 本页扫描结束。")
            unmatched_titles = []
            for nt in expected_titles:
                if not nt['matched']:
                    print(f"  [警告-未命中] 预期标题 '{nt['title']}' 未能在本页找到完美契合！(可能已被保护或属于错位引用)")
                    unmatched_titles.append(nt)
            
            if inject_title and unmatched_titles:
                inject_str_list = []
                for nt in unmatched_titles:
                    print(f"  [强制注入] 已激活强制补全机制，强行在顶部插入未匹配的标题: '{nt['title']}'")
                    inject_str_list.append(f"[SYS_TITLE]{nt['title']}[/SYS_TITLE]")
                    nt['matched'] = True
                result_lines = inject_str_list + result_lines

        return "\n".join(result_lines)

class RedXCheckBox(QtWidgets.QCheckBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(30, 30)
        self.setStyleSheet("""
            QCheckBox::indicator { width: 28px; height: 28px; border: 1px solid #aaa; border-radius: 4px; background: white; }
            QCheckBox::indicator:checked { background-color: #ff4d4f; border: 1px solid #d9363e; }
        """)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.isChecked():
            p = QtGui.QPainter(self)
            p.setRenderHint(QtGui.QPainter.Antialiasing)
            p.setPen(QtGui.QPen(QtGui.QColor("white"), 3))
            m = 8
            p.drawLine(m, m, self.width()-m, self.height()-m)
            p.drawLine(self.width()-m, m, m, self.height()-m)
            p.end()

class ClickableFilenameLabel(QtWidgets.QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self._full_text = text

    def mouseDoubleClickEvent(self, event):
        QtWidgets.QMessageBox.information(self, "完整文件名", self._full_text)
        super().mouseDoubleClickEvent(event)

# 指定导出的文件夹默认路径
DEFAULT_EXPORT_PATH = r""

class MainDialog(QtWidgets.QDialog):
    def __init__(self, bk):
        super().__init__()
        self.bk = bk
        self.converter = BBCodeConverter(bk)
        self.deleted_imgs = set()
        self.plugin_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.pref_file = os.path.join(self.plugin_dir, "bbcode_config.json")
        self.rule_file = os.path.join(self.plugin_dir, "rule.json")
        
        self.book_title = "Export"
        try:
            # 【修复】回归精准正则解析，规避 bs4 html.parser 对 XML 命名空间(dc:title)的吞标签问题
            opf = bk.get_opf()
            t = re.search(r'<(?:dc:)?title[^>]*id=["\']title["\'][^>]*>(.*?)</(?:dc:)?title>', opf, re.I|re.S)
            if not t: 
                t = re.search(r'<(?:dc:)?title[^>]*>(.*?)</(?:dc:)?title>', opf, re.I|re.S)
            if t: 
                raw_title = re.sub(r'<[^>]+>', '', t.group(1)).strip()
                self.book_title = re.sub(r'[\\/:*?"<>|]', '_', raw_title)
        except Exception as e: 
            print(f"[!] 书名读取异常: {e}")

        self.init_ui()

    def init_ui(self):
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self.setWindowTitle("EPUB → BBCode TXT (多核引擎预编译提速版)")
        self.resize(950, 850)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.stack = QtWidgets.QStackedWidget()
        self.layout.addWidget(self.stack)

        # --- 页面 1 ---
        self.page1 = QtWidgets.QWidget()
        p1_layout = QtWidgets.QVBoxLayout(self.page1)
        self.text_edit = QtWidgets.QTextEdit()
        tpl = f"{self.book_title}\n───────────────────────────\n[find]作者：[/find]\n───────────────────────────\n[b] 內容簡介 [/b]\n"
        if os.path.exists(self.pref_file):
            try:
                with open(self.pref_file, 'r', encoding='utf-8') as f: 
                    saved_tpl = json.load(f).get("template", tpl)
                    lines = saved_tpl.split('\n')
                    if lines: lines[0] = self.book_title
                    tpl = '\n'.join(lines)
            except: pass
        self.text_edit.setPlainText(tpl)
        p1_layout.addWidget(QtWidgets.QLabel("步骤 1: 制作信息配置"))
        p1_layout.addWidget(self.text_edit)
        
        b1 = QtWidgets.QHBoxLayout()
        sv = QtWidgets.QPushButton("保存模板"); sv.clicked.connect(self.save_tpl)
        rl = QtWidgets.QPushButton("打开/编辑 rule.json"); rl.clicked.connect(self.open_rule_json)
        fd = QtWidgets.QPushButton("查找 (Find)"); fd.clicked.connect(self.do_find)
        st = QtWidgets.QPushButton("下一步 (Next)"); st.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        b1.addWidget(sv); b1.addWidget(rl); b1.addStretch(); b1.addWidget(fd); b1.addWidget(st)
        p1_layout.addLayout(b1)
        self.stack.addWidget(self.page1)

        # --- 页面 2 ---
        self.page2 = QtWidgets.QWidget()
        p2_layout = QtWidgets.QVBoxLayout(self.page2)
        p2_layout.addWidget(QtWidgets.QLabel("步骤 2: 图床导入 (点击左侧框出现红色 × 以删除图片)"))
        
        self.scroll = QtWidgets.QScrollArea()
        self.scroll_content = QtWidgets.QWidget()
        self.img_layout = QtWidgets.QVBoxLayout(self.scroll_content)
        self.img_items = []

        try:
            img_list = []
            for info in self.bk.image_iter():
                iid = info[0]
                href = info[1] if len(info) > 1 and info[1] else ""
                name = os.path.basename(href or iid)
                img_list.append((name, iid))
            
            img_list.sort(key=lambda x: x[0])

            for name, iid in img_list:
                row = QtWidgets.QHBoxLayout()
                row.setContentsMargins(5, 5, 5, 5)
                
                chk = RedXCheckBox()
                
                img_lbl = QtWidgets.QLabel()
                img_lbl.setFixedSize(80, 80)
                img_lbl.setStyleSheet("border: 1px solid #ddd; background: #f9f9f9;")
                img_lbl.setAlignment(QtCore.Qt.AlignCenter)
                try:
                    pix = QtGui.QPixmap()
                    pix.loadFromData(self.bk.readfile(iid))
                    img_lbl.setPixmap(pix.scaled(80, 80, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
                except: 
                    img_lbl.setText("ERR")
                
                lbl_name = ClickableFilenameLabel(name)
                lbl_name.setFixedWidth(150)
                lbl_name.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                
                edit = QtWidgets.QLineEdit()
                edit.setPlaceholderText("粘贴 URL...")
                edit.setFixedHeight(30) 
                
                row.addWidget(chk)
                row.addWidget(img_lbl)
                row.addWidget(lbl_name)
                row.addWidget(edit, 1)
                
                self.img_layout.addLayout(row)
                self.img_items.append({'name': name, 'checkbox': chk, 'edit': edit})
        except: pass

        self.scroll.setWidget(self.scroll_content)
        self.scroll.setWidgetResizable(True)
        p2_layout.addWidget(self.scroll)

        b2 = QtWidgets.QHBoxLayout()
        btn_clear = QtWidgets.QPushButton("清除链接"); btn_clear.clicked.connect(self.do_clear_urls)
        pt = QtWidgets.QPushButton("批量添加"); pt.clicked.connect(self.do_paste)
        
        self.chk_inject_title = QtWidgets.QCheckBox("注入标题")
        self.chk_inject_title.setChecked(False)

        btn_prev = QtWidgets.QPushButton("上一步 (Prev)"); btn_prev.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.cf = QtWidgets.QPushButton("确认转换"); self.cf.clicked.connect(self.do_convert)
        
        b2.addWidget(btn_clear)
        b2.addWidget(pt)
        b2.addWidget(self.chk_inject_title)
        b2.addStretch()
        b2.addWidget(btn_prev)
        b2.addWidget(self.cf)
        
        p2_layout.addLayout(b2)
        self.stack.addWidget(self.page2)

    def open_rule_json(self):
        if not os.path.exists(self.rule_file):
            try:
                default_content = { "全角转半角": {}, "符号转换": {}, "标点修正": {} }
                with open(self.rule_file, 'w', encoding='utf-8') as f:
                    json.dump(default_content, f, ensure_ascii=False, indent=2)
            except: pass
        import platform, subprocess
        try:
            if platform.system() == 'Windows': os.startfile(self.rule_file)
            elif platform.system() == 'Darwin': subprocess.call(['open', self.rule_file])
            else: subprocess.call(['xdg-open', self.rule_file])
        except: QtWidgets.QMessageBox.warning(self, "提示", f"请手动打开:\n{self.rule_file}")

    def save_tpl(self):
        try:
            with open(self.pref_file, 'w', encoding='utf-8') as f:
                json.dump({"template": self.text_edit.toPlainText()}, f, ensure_ascii=False)
            QtWidgets.QMessageBox.information(self, "成功", "模板已保存。")
        except: pass

    def do_find(self):
        raw = self.text_edit.toPlainText()
        txt = ""
        for info in self.bk.text_iter(): 
            txt += re.sub(r'<[^>]+>', '', self.bk.readfile(info[0]))
        def repl(m):
            k = m.group(1)
            r = re.search(re.escape(k) + r'(.*?)\n', txt)
            return f"{k}{r.group(1).strip()}" if r else f"{k}"
        self.text_edit.setPlainText(re.sub(r'\[find\](.*?)\[/find\]', repl, raw))

    def do_clear_urls(self):
        reply = QtWidgets.QMessageBox.question(self, '二次确认', '确定要清除所有已填写的链接吗？\n该操作不可撤销。', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            for it in self.img_items: it['edit'].clear()

    def do_paste(self):
        reply = QtWidgets.QMessageBox.question(self, '二次确认', '确定要使用剪贴板内容批量覆盖当前全部链接吗？\n（注：已打勾删除的图片将被自动跳过）', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            cb = QtWidgets.QApplication.clipboard().text().strip('\r\n').split('\n')
            cb = [l.strip() for l in cb]
            
            url_idx = 0
            for it in self.img_items:
                if it['checkbox'].isChecked():
                    continue
                if url_idx < len(cb):
                    it['edit'].setText(cb[url_idx])
                    url_idx += 1

    def do_convert(self):
        self.cf.setEnabled(False)
        self.deleted_imgs = {it['name'] for it in self.img_items if it['checkbox'].isChecked()}
        mapping = {it['name']: f"[img]{it['edit'].text().strip() if it['edit'].text().strip() else it['name']}[/img]" 
                   for it in self.img_items if it['name'] not in self.deleted_imgs}
        
        current_rules = {}
        if os.path.exists(self.rule_file):
            try:
                with open(self.rule_file, 'r', encoding='utf-8-sig') as f:
                    current_rules = json.load(f)
            except Exception as e:
                print(f"[!] rule.json 加载失败: {e}")

        prog = None
        try:
            spine_list = list(self.bk.spine_iter())
            prog = QtWidgets.QProgressDialog("执行单向合并重构引擎...", "取消", 0, len(spine_list), self)
            prog.setWindowModality(QtCore.Qt.WindowModal)
            prog.show()

            self.converter.pre_scan()
            info = self.text_edit.toPlainText().strip()
            info = re.sub(r'(\[b\]\s*內容簡介\s*\[/b\]\n?)(.*)', r'\n[center]\1\2[/center]\n', info, flags=re.S|re.I)
            
            bodies = []
            print("\n" + "="*50 + "\n【第二步：按照 SPINE 书脊顺序合并正文，并触发动态匹配】\n" + "="*50)

            manifest_map = {}
            for m_uid, m_href, m_mime in self.bk.manifest_iter():
                manifest_map[m_uid] = m_href
                
            inject_title_flag = self.chk_inject_title.isChecked()

            for i, spine_info in enumerate(spine_list):
                uid = spine_info[0]
                href = manifest_map.get(uid, "")
                filename = os.path.basename(href.split('#')[0] if href else uid)
                
                if prog.wasCanceled(): return
                prog.setValue(i); QtWidgets.QApplication.processEvents()

                if filename == self.converter.nav_filename:
                    print(f"--- [物理剥离] 跳过合并官方导航文件: {filename} ---")
                    continue
                if filename in self.converter.skip_files:
                    print(f"--- [物理剥离] 强制丢弃受保护的特权页面 (如版权页): {filename} ---")
                    continue

                html_raw = self.bk.readfile(uid)
                body_match = re.search(r'<body[^>]*>(.*?)</body>', html_raw, flags=re.S|re.I)
                if body_match:
                    page_text = self.converter.clean_and_convert(body_match.group(1), href, mapping, self.deleted_imgs, current_idx=i, inject_title=inject_title_flag)
                    
                    if filename in self.converter.toc_files:
                        print(f"--- [特权处理] 发现目录页: {filename}，已为其整页首尾包裹 [center] 标签 ---")
                        page_text = f"[center]{page_text.strip()}[/center]"
                        
                    bodies.append(page_text)

            info = info.strip()
            lines = info.split('\n')
            if lines and lines[0] != self.book_title:
                info = f"{self.book_title}\n\n{info}"

            final = info + "\n" + "\n".join(bodies)

            print("\n[*] 全局重构完毕，正在进行最终空间修补与 rule.json 应用...")

            final = final.replace('\r\n', '\n').replace('\r', '\n')
            
            if current_rules:
                for cat, rs in current_rules.items():
                    for pat, r_text in rs.items():
                        try:
                            safe_pat = pat.replace('\\\\', '\\')
                            final = re.sub(safe_pat, r_text, final)
                        except Exception as e: 
                            print(f"[!] 正则 '{pat}' 发生错误: {e}")

            final = re.sub(r'(?m)^[ \t\u3000]+$\n?', '', final)

            junk_line = r'\n(?:[ \t\u3000]|\[SYS_BR_SPACE\])*(?=\n|$)'

            pattern_img = r'(?:' + junk_line + r')*\n*[ \t\u3000]*(\[img\].*?\[/img\])[ \t\u3000]*(?=\n|$)(?:' + junk_line + r')*'
            final = re.sub(pattern_img, r'\n\n\1\n\n', final)
            
            pattern_title = r'(?:' + junk_line + r')*\n*[ \t\u3000]*\[SYS_TITLE\](.*?)\[/SYS_TITLE\][ \t\u3000]*(?=\n|$)(?:' + junk_line + r')*'
            final = re.sub(pattern_title, r'\n\n[center][b]\1[/b][/center]\n\n', final)

            final = re.sub(r'\[segmentation\](.*?)\[/segmentation\]', r'[center][b]\1[/b][/center]', final)

            final = re.sub(r'\n{3,}', '\n\n', final)

            final = final.replace('[SYS_BR_SPACE]', '　')

            final = re.sub(r'\n*\[SYS_HR_MARKER\]\n*', '\n[hr]\n', final)
            final = final.replace('[SYS_HR_MARKER]', '[hr]') 

            while '[img][img]' in final:
                final = final.replace('[img][img]', '[img]')
            while '[/img][/img]' in final:
                final = final.replace('[/img][/img]', '[/img]')

            print("\n【✅ 转换彻底完成】控制台报告已输出完毕。")
            
            safe_filename = re.sub(r'[\\/:*?"<>|]', '_', self.book_title)
            
            default_save_path = ""
            if DEFAULT_EXPORT_PATH and os.path.isdir(DEFAULT_EXPORT_PATH):
                default_save_path = os.path.join(DEFAULT_EXPORT_PATH, f"{safe_filename}.txt")
            else:
                default_save_path = f"{safe_filename}.txt"

            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出最终结果", default_save_path, "TXT 文件 (*.txt)")
            if path:
                with open(path, 'w', encoding='utf-8') as f: f.write(final.strip() + "\n")
                self.accept()
        except Exception as err:
            err_details = traceback.format_exc()
            print("\n【致命崩溃 Traceback】\n" + err_details)
            try:
                with open(os.path.join(self.plugin_dir, "error_log.txt"), "w", encoding="utf-8") as f:
                    f.write("\n".join(console_logs))
            except: pass
            QtWidgets.QMessageBox.critical(self, "程序异常", f"遇到致命崩溃：\n\n{str(err)}\n\n(完整堆栈已保存至 error_log.txt)")
        finally:
            try:
                with open(os.path.join(self.plugin_dir, "error_log.txt"), "w", encoding="utf-8") as f:
                    f.write("\n".join(console_logs))
            except: pass
            self.cf.setEnabled(True)
            if prog: prog.close()

def run(bk):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    dialog = MainDialog(bk)
    if hasattr(dialog, 'exec'):
        dialog.exec()
    else:
        dialog.exec_()
    return 0
