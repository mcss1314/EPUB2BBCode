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

# ========================================================
# 修复跨设备/不同Sigil版本导致的 Qt platform plugin 冲突问题
# 强制清除Sigil传递的Qt环境变量，让PyQt/PySide使用自带的platform plugins
# ========================================================
for _env_key in ['QT_PLUGIN_PATH', 'QT_QPA_PLATFORM_PLUGIN_PATH']:
    if _env_key in os.environ:
        del os.environ[_env_key]

# 动态加载 Qt 库
e = os.environ.get('SIGIL_QT_RUNTIME_VERSION', '6.5.2')
SIGIL_QT_MAJOR_VERSION = tuple(map(int, (e.split("."))))[0]
if SIGIL_QT_MAJOR_VERSION == 6:
    from PySide6 import QtWidgets, QtCore, QtGui
elif SIGIL_QT_MAJOR_VERSION == 5:
    from PyQt5 import QtWidgets, QtCore, QtGui

class BBCodeConverter:
    def __init__(self, bk):
        self.bk = bk
        self.footnote_map = {}
        self.nav_titles = []
        self.nav_uid = ""
        self.nav_filename = ""
        self.skip_files = set()
        self.toc_files = {}  # 独立记录目录页及其名称
        self.spine_map = {}  # 记录文件对应的 OPF Spine 顺序索引，构建单向状态机防回溯

    def add_nav_title(self, title, href_full):
        if not title or not href_full: return
        # 严格遵守要求：只定位xhtml，不识别后面的#id，避免匹配错误
        filename = os.path.basename(href_full.split('#')[0])
        title = title.strip(' \t\n\r\u3000')
        
        # 【特权】封面处理：仅忽略标题匹配，保留页面正常内容处理
        if title in ["書封", "书封", "封面", "作者頁", "作者页", "書名頁", "书名页", "內彩", "封底"]:
            print(f"  -> [识别-忽略] 发现封面标题: '{title}'，不将其纳入匹配库（但保留页面正文读取）")
            return
            
        # 记录标题应该存在的目标文件的序号 (防回溯核心)
        target_idx = self.spine_map.get(filename, 9999)
        
        # 【特权】目录与版权页
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
            # 构建 manifest 映射以安全获取 href
            manifest_map = {}
            for m_uid, m_href, m_mime in self.bk.manifest_iter():
                manifest_map[m_uid] = m_href

            # 预先扫描书脊，构建防回溯时间轴
            for idx, spine_info in enumerate(self.bk.spine_iter()):
                uid = spine_info[0]
                href = manifest_map.get(uid, "")
                if href:
                    self.spine_map[os.path.basename(href.split('#')[0])] = idx

            opf = self.bk.get_opf()
            m_nav = re.search(r'<item[^>]*id=["\']([^"\']+)["\'][^>]*properties=["\'][^"\']*nav[^"\']*["\']', opf, re.I)
            if not m_nav:
                m_nav = re.search(r'<item[^>]*properties=["\'][^"\']*nav[^"\']*["\'][^>]*id=["\']([^"\']+)["\']', opf, re.I)
            m_toc = re.search(r'<spine[^>]*toc=["\']([^"\']+)["\']', opf, re.I)
            
            self.nav_uid = m_nav.group(1) if m_nav else (m_toc.group(1) if m_toc else "toc")
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
                print("[*] 开始从导航中提取预设标题...")
                
                # 【防护增强】仅锁定 epub:type="toc" 的范围，彻底抛弃 landmarks 等干扰元素
                toc_block = re.search(r'(<nav[^>]*epub:type=["\']toc["\'][^>]*>.*?</nav>)', nav_html, flags=re.I|re.S)
                if toc_block:
                    nav_html = toc_block.group(1)
                    print("  -> [防护] 成功隔离 <nav epub:type=\"toc\"> 核心区块，有效屏蔽 landmarks。")
                    
                if self.nav_filename.lower().endswith('.ncx'):
                    for m in re.finditer(r'<navPoint[^>]*>.*?<text[^>]*>(.*?)</text>.*?<content[^>]*src=["\']([^"\']+)["\']', nav_html, flags=re.S|re.I):
                        self.add_nav_title(re.sub(r'<[^>]+>', '', m.group(1)), m.group(2))
                else:
                    for m in re.finditer(r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', nav_html, flags=re.I|re.S):
                        self.add_nav_title(re.sub(r'<[^>]+>', '', m.group(2)), m.group(1))
            except Exception as e:
                print(f"[!] 导航解析发生异常: {e}")

        # 全局备用脚注（如果当页找不到，可以来这里找）
        try:
            print("[*] 开始提取全书备用脚注...")
            for text_info in self.bk.text_iter():
                uid, href = text_info[0], text_info[1] if len(text_info)>1 else ""
                filename = os.path.basename(href or "")
                if uid == self.nav_uid or filename == self.nav_filename or filename in self.skip_files or filename in self.toc_files: 
                    continue
                
                content = self.bk.readfile(uid)
                for aside_m in re.finditer(r'<aside[^>]*>(.*?)</aside>', content, re.S|re.I):
                    for li_m in re.finditer(r'<li[^>]*id=["\']([^"\']+)["\'][^>]*>(.*?)</li>', aside_m.group(1), re.S|re.I):
                        txt = re.sub(r'<[^>]+>', '', li_m.group(2)).strip(' \t\n\r\u3000')
                        if txt: self.footnote_map[li_m.group(1)] = txt
                for fn_m in re.finditer(r'<[^>]+class=["\'][^"\']*footnote[^"\']*["\'][^>]*id=["\']([^"\']+)["\'][^>]*>(.*?)</(?:p|div|span|li)>', content, re.S|re.I):
                    txt = re.sub(r'<[^>]+>', '', fn_m.group(2)).strip(' \t\n\r\u3000')
                    if txt: self.footnote_map[fn_m.group(1)] = txt
            print(f"[*] 共提取到 {len(self.footnote_map)} 个备用脚注。")
        except Exception as e:
            print(f"[!] 备用脚注提取发生异常: {e}")

    def clean_and_convert(self, html, href, img_map, deleted_imgs, current_idx, inject_title=False):
        href = href or ""
        filename = os.path.basename(href.split('#')[0])
        
        # 激活条件：标题所在文件一致，且尚未被匹配，且当前文件 OPF Index 必须大等于其目标 Index
        expected_titles = [nt for nt in self.nav_titles if nt['filename'] == filename and not nt['matched'] and current_idx >= nt['target_idx']]

        print(f"\n--- 【文件处理开始】: {filename} (当页绝对索引: {current_idx}) ---")
        if expected_titles:
            print(f"  [监测-标题池] 本页准许并激活匹配 {len(expected_titles)} 个标题。")

        # 侦测表格标签并在控制台警告 (不作处理)
        if re.search(r'<table\b', html, flags=re.I):
            print(f"  [警告] 侦测到表格标签 <table>，当前版本不支持表格转换，请注意输出排版。")

        # ========================================================
        # 1. 基础间隙清理 (不碰全角空格)
        # ========================================================
        html = re.sub(r'>\s*\n\s*<', '><', html)

        # ========================================================
        # 2. 注释（脚注）精准拉取与销毁系统 (重写变量轮询判断先后顺序)
        # ========================================================
        print(f"  [处理-注释] 开始扫描当页脚注定义...")
        local_footnotes = {}
        del_blocks = []
        
        # 第一步：提取所有包含 注/註/* 的 <a> 标签
        a_tags = re.finditer(r'<a[^>]*href=["\']#([^"\']+)["\'][^>]*>(.*?)</a>', html, flags=re.I|re.S)
        
        for m in a_tags:
            link_pos = m.start()
            target_id = m.group(1)
            a_text_raw = re.sub(r'<[^>]+>', '', m.group(2)).strip(' \t\n\r\u3000[]()【】')
            
            # 判断是否是注释链接
            if '注' in a_text_raw or '註' in a_text_raw or '*' in a_text_raw or a_text_raw.isdigit():
                # 轮询查找此 target_id
                target_str1 = f'id="{target_id}"'
                target_str2 = f"id='{target_id}'"
                
                idx = html.find(target_str1)
                if idx == -1: idx = html.find(target_str2)
                
                # 找到了，并且这个 id 出现在当前链接的后面（即"第二个"，正文在前，注释定义在后）
                if idx != -1 and idx > link_pos:
                    # 向左寻找最近的块级包裹
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
                            
                            # 提取纯文本
                            txt = re.sub(r'<[^>]+>', '', block_html).strip(' \t\n\r\u3000')
                            txt = re.sub(r'\s+', ' ', txt)
                            txt = re.sub(r'^[\^↵↑*※]\s*', '', txt)
                            txt = re.sub(r'^返回正文\s*', '', txt)
                            txt = re.sub(r'^' + re.escape(a_text_raw) + r'[:：\s]*', '', txt)
                            
                            local_footnotes[target_id] = {'text': txt}
                            
                            # 标记待删除（确保不重复）
                            if block_html not in del_blocks:
                                del_blocks.append(block_html)

        # 第二步：将所有被捕获的注释块打上临时标记
        for block in del_blocks:
            if block in html:
                html = html.replace(block, f'[SYS_DEL_FOOTNOTE]{block}[/SYS_DEL_FOOTNOTE]')
        if del_blocks:
            print(f"  [处理-注释销毁准备] 已为 {len(del_blocks)} 个底部注释区块添加删除标识符。")

        # 第三步：执行文中链接的替换注入
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
                
        html, fn_count = re.subn(r'<a[^>]*href=["\']#([^"\']+)["\'][^>]*>.*?</a>', footnote_link_replacer, html, flags=re.I|re.S)
        if fn_count > 0:
            print(f"  [处理-注释] 发现并内联了 {fn_count} 处脚注引用。")

        # 第四步：注释处理结束后，统一删除被标记的注释块
        html, del_count = re.subn(r'\[SYS_DEL_FOOTNOTE\].*?\[/SYS_DEL_FOOTNOTE\]', '', html, flags=re.I|re.S)
        if del_count > 0:
            print(f"  [处理-注释销毁执行] 成功清除了 {del_count} 个被标记的底部注释区块。")

        # ========================================================
        # 3. 清理空标签 (不清理被保护的内容)
        # ========================================================
        while True:
            new_html = re.sub(r'<([a-z0-9]+)[^>]*>\s*</\1>', '', html, flags=re.I)
            if new_html == html: break
            html = new_html

        # ========================================================
        # 4. 保护分行符 (start-6em)
        # ========================================================
        def mask_div(m):
            txt = re.sub(r'<[^>]+>', '', m.group(1)).strip(' \t\n\r\u3000')
            txt = re.sub(r'\s+', ' ', txt) # 强制扁平化，防止内部污染
            # 必须包裹在原有的 <p> 标签内，确保它在后续剥离标签时能正常换行
            return f'<p class="start-6em">[segmentation]{txt}[/segmentation]</p>'
        html, d_count = re.subn(r'<p[^>]+class=["\'][^"\']*start-6em[^"\']*["\'][^>]*>(.*?)</p>', mask_div, html, flags=re.S | re.I)
        if d_count > 0: print(f"  [保护-分行符] 成功掩护了 {d_count} 个 start-6em 标签。")

        # ========================================================
        # 5. Ruby 处理 (严格保留正文空格)
        # ========================================================
        html = re.sub(r'</rt>\s+</ruby>', '</rt></ruby>', html, flags=re.I)
        def ruby_repl(m):
            inner = m.group(1)
            rts = "".join([re.sub(r'<[^>]+>', '', c) for c in re.findall(r'<rt[^>]*>(.*?)</rt>', inner, flags=re.I|re.S)])
            base = re.sub(r'<(?:rp|rt)[^>]*>.*?</(?:rp|rt)>', '', inner, flags=re.I|re.S)
            base = re.sub(r'<[^>]+>', '', base).replace('\n', '')
            return f"[ruby={rts}]{base}[/ruby]"
        html, r_count = re.subn(r'<ruby[^>]*>(.*?)</ruby>', ruby_repl, html, flags=re.I | re.S)
        if r_count > 0: print(f"  [处理-Ruby] 成功转换了 {r_count} 处 Ruby 注音。")

        # ========================================================
        # 6. 图片与基础替换 (含 <br/> 强力保护机制)
        # ========================================================
        # 【完美修复】将 <p> 标签内的 <br/> 预先转化为系统代号，完美避开后续任何"压缩"与"清空"
        html, b_count = re.subn(r'<p[^>]*>(.*?)</p>', lambda m: "<p>" + re.sub(r'<br\s*/?>', '[SYS_BR_SPACE]', m.group(1), flags=re.I) + "</p>", html, flags=re.S|re.I)
        
        def img_repl(m):
            src = os.path.basename(m.group(1))
            return "" if src in deleted_imgs else img_map.get(src, f"__IMG_MARKER__{src}__")
        html, i_count = re.subn(r'<(?:img|image)[^>]+(?:src|href)=["\']([^"\']+)["\'][^>]*>', img_repl, html, flags=re.I)
        if i_count > 0: print(f"  [处理-图片] 成功映射了 {i_count} 张图片。")

        # 【深度递归修复】解决标签嵌套导致提前闭合的恶性Bug
        # 从最内层标签开始向外逐层剥离或替换，完美支持 <span class="gfont">1<span>2</span>3</span> 及其内层样式
        while True:
            # (?![^>]*?<(?:span|b|strong|i|em|s)\b) 确保我们定位到的是没有内部重叠标签的最内层标签
            m = re.search(r'<(span|b|strong|i|em|s)\b([^>]*)>((?:(?!<(?:span|b|strong|i|em|s)\b).)*?)</\1>', html, flags=re.S|re.I)
            if not m: break
            
            tag = m.group(1).lower()
            attrs = m.group(2).lower()
            inner = m.group(3)
            
            if tag in ('b', 'strong'):
                repl = f"[b]{inner}[/b]"
            elif tag in ('i', 'em'):
                repl = f"[i]{inner}[/i]"
            elif tag == 's':
                repl = f"[s]{inner}[/s]"
            elif tag == 'span':
                if 'bold' in attrs or 'gfont' in attrs:
                    repl = f"[b]{inner}[/b]"
                elif 'italic' in attrs:
                    repl = f"[i]{inner}[/i]"
                else:
                    repl = inner  # 剥离无意义外壳，保留内容供外层处理
            else:
                repl = inner
            
            html = html[:m.start()] + repl + html[m.end():]

        # ========================================================
        # 7. 剥离块级标签，转化为真实行
        # ========================================================
        html = re.sub(r'</p>|</div>|</li>|</h[1-6]>|<br\s*/?>', '\n', html, flags=re.I)
        text = re.sub(r'<[^>]+>', '', html)

        # ========================================================
        # 8. 单向防回溯智能标题侦测与捕获
        # ========================================================
        raw_lines = text.split('\n')
        marked_lines = []

        print(f"  [处理-标题比对] 开始进入核心按行扫描...")
        for line_num, line in enumerate(raw_lines, 1):
            stripped_line = line.strip(' \t\r\n\u3000')
            if not stripped_line:
                marked_lines.append(line)
                continue

            # 遇到受沙盒保护的分行符，直接放行，杜绝被误判为标题
            if '[segmentation]' in stripped_line:
                marked_lines.append(line)
                continue

            # 【修复点】忽略（剔除）本行内的图片标记，仅提取纯净文字用于标题比对
            temp_line = re.sub(r'\[img\].*?\[/img\]', '', stripped_line, flags=re.I)
            temp_line = re.sub(r'__IMG_MARKER__.*?__', '', temp_line)

            # 获取完全纯粹的文字用于兜底比对（例如剔除 [i]、[/i] 等伪装）
            pure_text = re.sub(r'\[/?[a-z=0-9]+\]', '', temp_line).strip(' \t\r\n\u3000')
            norm_line = re.sub(r'\s+', '', pure_text)
            match_found = False

            if norm_line and expected_titles:
                for nt in expected_titles:
                    if not nt['matched']:
                        exp_norm = re.sub(r'\s+', '', nt['title'])
                        
                        # 【日志增强】对有可能包含目标字样的行输出跟踪信息，方便定位
                        if len(norm_line) > 1 and (norm_line in exp_norm or exp_norm in norm_line):
                             print(f"    [分析追踪 | 行 {line_num:03d}] 探测到包含标题核心字的行。目标:'{exp_norm}' | 剥离后提取为:'{norm_line}'")
                        
                        if norm_line == exp_norm:
                            print(f"    ✅ [命中-成功锁定首出位 | 行 {line_num:03d}] 原文提取的 '{pure_text}' == 导航标题 '{nt['title']}'。已加锁。")
                            # 采用免伤的方括号封闭式沙盒标签
                            marked_lines.append(f"[SYS_TITLE]{nt['title']}[/SYS_TITLE]")
                            nt['matched'] = True
                            match_found = True
                            break

            if not match_found:
                marked_lines.append(line)

        # 章节战报结算与强制标题注入逻辑
        result_lines = marked_lines
        if expected_titles:
            print(f"  [章节战报结算] 本页扫描结束。")
            unmatched_titles = []
            for nt in expected_titles:
                if not nt['matched']:
                    print(f"  [警告-未命中] 预期标题 '{nt['title']}' 未能在本页找到完美契合！(可能已被保护或属于错位引用)")
                    unmatched_titles.append(nt)
            
            # 如果开启了注入开关，且该章节有未被匹配上的期待标题，则强制插入文档最顶端
            if inject_title and unmatched_titles:
                inject_str_list = []
                for nt in unmatched_titles:
                    print(f"  [强制注入] 已激活强制补全机制，强行在顶部插入未匹配的标题: '{nt['title']}'")
                    inject_str_list.append(f"[SYS_TITLE]{nt['title']}[/SYS_TITLE]")
                    nt['matched'] = True  # 标记为已解决，防止串流
                # 将强制注入的标题加在原始页面的最上面
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

# 指定导出的文件夹默认路径（例如 r"C:\Downloads"）。留空即 r"" 也会自动回退系统默认，不会报错。
DEFAULT_EXPORT_PATH = r""

class MainDialog(QtWidgets.QDialog):
    def __init__(self, bk):
        super().__init__()
        self.bk = bk
        self.converter = BBCodeConverter(bk)
        self.deleted_imgs = set()
        self.plugin_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 【关键修改】强行将配置文件和 rule.json 一样，绑定在插件源码所在的安装绝对路径下！
        self.pref_file = os.path.join(self.plugin_dir, "bbcode_config.json")
        self.rule_file = os.path.join(self.plugin_dir, "rule.json")
        
        self.book_title = "Export"
        try:
            # 【关键修复】精准定位读取OPF的 <dc:title id="title"> 标签
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
        # 移除窗口右上角的上下文帮助问号按钮
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self.setWindowTitle("EPUB → BBCode TXT (空间矩阵引擎·原生界面日志强化版)")
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
                    # 保证用户加载历史配置时，第一行也会强制替换为当前最新书名
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
            for info in self.bk.image_iter():
                iid = info[0]
                href = info[1] if len(info) > 1 and info[1] else ""
                name = os.path.basename(href or iid)
                
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
                except: img_lbl.setText("ERR")
                
                # 【修改】文件名称宽度改为目前的四分之三 (约176px)，取消换行，双击弹窗
                lbl_name = ClickableFilenameLabel(name)
                lbl_name.setFixedWidth(176)
                lbl_name.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                
                edit = QtWidgets.QLineEdit()
                edit.setPlaceholderText("粘贴 URL...")
                edit.setFixedHeight(30) 
                
                # 【修改】调换顺序：选项 -> 图片 -> 文件名 -> 链接输入框
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
        
        # 新增注入标题开关
        self.chk_inject_title = QtWidgets.QCheckBox("注入标题")
        self.chk_inject_title.setChecked(False)

        # 新增上一步功能，返回时不清除页面状态
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
                    continue # 跳过删除行，不消耗剪贴板内容
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
            # 严格依据 spine_iter 的顺序进行合并
            spine_list = list(self.bk.spine_iter())
            prog = QtWidgets.QProgressDialog("执行单向合并重构引擎...", "取消", 0, len(spine_list), self)
            prog.setWindowModality(QtCore.Qt.WindowModal)
            prog.show()

            self.converter.pre_scan()
            info = self.text_edit.toPlainText().strip()
            info = re.sub(r'(\[b\]\s*內容簡介\s*\[/b\]\n?)(.*)', r'\n[center]\1\2[/center]\n', info, flags=re.S|re.I)
            
            bodies = []
            print("\n" + "="*50 + "\n【第二步：按照 SPINE 书脊顺序合并正文，并触发动态匹配】\n" + "="*50)

            # 构建全局 manifest 映射，彻底解决 spine_iter 提取 href 错乱（变成 "yes"）的致命问题
            manifest_map = {}
            for m_uid, m_href, m_mime in self.bk.manifest_iter():
                manifest_map[m_uid] = m_href
                
            # 抓取是否需要开启“强制注入标题”功能
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
                    # 传入当前索引 i，作为防回溯的校验基准，并携带强制注入标志
                    page_text = self.converter.clean_and_convert(body_match.group(1), href, mapping, self.deleted_imgs, current_idx=i, inject_title=inject_title_flag)
                    
                    # 【特权】目录页：处理完整页内容后，在首行最左侧和末行最右侧添加 [center] 标签
                    if filename in self.converter.toc_files:
                        print(f"--- [特权处理] 发现目录页: {filename}，已为其整页首尾包裹 [center] 标签 ---")
                        page_text = f"[center]{page_text.strip()}[/center]"
                        
                    bodies.append(page_text)

            # 【最后防线】绝对确保最终文本第一行是读取的书名
            info = info.strip()
            lines = info.split('\n')
            if lines and lines[0] != self.book_title:
                info = f"{self.book_title}\n\n{info}"

            final = info + "\n" + "\n".join(bodies)

            # ========================================================
            # 最终定向排版引擎 (绝对执行物理级别的空间坍缩与修复)
            # ========================================================
            print("\n[*] 全局重构完毕，正在进行最终空间修补与 rule.json 应用...")

            # 【前置防御】剿灭隐形的 \r 字符，防止它破坏换行与空格行正则匹配！
            final = final.replace('\r\n', '\n').replace('\r', '\n')
            
            if current_rules:
                for cat, rs in current_rules.items():
                    for pat, r_text in rs.items():
                        try:
                            safe_pat = pat.replace('\\\\', '\\')
                            final = re.sub(safe_pat, r_text, final)
                        except Exception as e: 
                            print(f"[!] 正则 '{pat}' 发生错误: {e}")

            # 【关键修复】使用多行正则，彻底删除仅由全角/半角空格组成的行
            # 加上 $ 严格限定必须一直到行尾全是空格，绝对防误伤正文的首行缩进！
            final = re.sub(r'(?m)^[ \t\u3000]+$\n?', '', final)

            # 【图床与标题的统一终极换行】
            # 采用全新的统一排版约束，完全废弃“标题下方无空行”的设定，让两者排版表现绝对一致：上下均唯一保留一个标准空行。
            junk_line = r'\n(?:[ \t\u3000]|\[SYS_BR_SPACE\])*(?=\n|$)'

            # 处理图片换行
            pattern_img = r'(?:' + junk_line + r')*\n*[ \t\u3000]*(\[img\].*?\[/img\])[ \t\u3000]*(?=\n|$)(?:' + junk_line + r')*'
            final = re.sub(pattern_img, r'\n\n\1\n\n', final)
            
            # 处理标题换行（现在和图片完全一致，替换尾部为 \n\n 以实现双向物理悬空排版）
            pattern_title = r'(?:' + junk_line + r')*\n*[ \t\u3000]*\[SYS_TITLE\](.*?)\[/SYS_TITLE\][ \t\u3000]*(?=\n|$)(?:' + junk_line + r')*'
            final = re.sub(pattern_title, r'\n\n[center][b]\1[/b][/center]\n\n', final)

            # 【分行符还原】完全按照您的要求，通过 [segmentation] 安全过渡后替换为最终 [center][b][/b][/center]
            final = re.sub(r'\[segmentation\](.*?)\[/segmentation\]', r'[center][b]\1[/b][/center]', final)

            # 【仅压缩极端连续空行，防误伤纯全角空格的独立行】
            final = re.sub(r'\n{3,}', '\n\n', final)

            # 【恢复保护】将被保护的 <br/> 还原为全角空格 (放在最后，避免被上方引擎当做无用空行误杀)
            final = final.replace('[SYS_BR_SPACE]', '　')

            print("\n【✅ 转换彻底完成】控制台报告已输出完毕。")
            
            # 导出文件名安全处理
            safe_filename = re.sub(r'[\\/:*?"<>|]', '_', self.book_title)
            
            # 【智能容错判定】留空或路径不存在时，完美降级回退到系统默认当前行为
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