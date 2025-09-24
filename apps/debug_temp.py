import os
import html
from collections import defaultdict
import io
import base64
from PIL import Image, ImageDraw
import fitz

class MockFileScannerApp:
    def _get_display_path(self, file_info):
        scan_root = file_info.get('scan_root', '')
        file_dir = file_info['path']
        if not scan_root or os.path.normpath(file_dir) == os.path.normpath(scan_root):
            return "."
        return os.path.relpath(file_dir, scan_root)

    def _create_placeholder_image(self, size, text):
        img = Image.new('RGB', size, color = (200, 200, 200))
        d = ImageDraw.Draw(img)
        d.text((10,10), text, fill=(0,0,0))
        return img

    def _generate_html_content(self, pages_to_export):
        grouped_pages = defaultdict(list)
        selected_pages_keys = {(os.path.join(p['file_info']['path'], p['file_info']['filename']), p['page_num']) for p in pages_to_export}
        
        for page in pages_to_export:
             full_path = os.path.join(page['file_info']['path'], page['file_info']['filename'])
             if (full_path, page['page_num']) in selected_pages_keys:
                grouped_pages[page['file_info']['scan_root']].append(page)

        file_colors = ['#DB4437', '#4285F4', '#F4B400', '#0F9D58', '#AB47BC', '#E91E63', '#9C27B0', '#673AB7', '#009688']
        level_colors = ['#4a789c', '#8b5f99', '#a3666c', '#948356', '#6e8b61']
        unique_file_paths = sorted({os.path.join(p['file_info']['path'], p['file_info']['filename']) for p in pages_to_export})
        file_path_to_color = {path: file_colors[i % len(file_colors)] for i, path in enumerate(unique_file_paths)}

        unique_folder_paths = sorted(grouped_pages.keys())
        folder_path_to_color = {path: file_colors[i % len(file_colors)] for i, path in enumerate(unique_folder_paths)}
        
        source_html = ""
        sorted_grouped_pages = sorted(grouped_pages.items(), key=lambda item: item[0])

        for folder_idx, (folder, pages) in enumerate(sorted_grouped_pages):
            display_folder = folder
            current_folder_color = folder_path_to_color.get(folder, '#808080')
            num_files = len({(p['file_info']['path'], p['file_info']['filename']) for p in pages})
            total_pages = len(pages)
            
            total_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0) for p in pages)
            total_trim_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('trim_area_sqm', p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0)) for p in pages)
            
            folder_stats_text = f"File: {num_files} | Pagine: {total_pages} | Area: {total_sqm:.2f} m²"
            if abs(total_sqm - total_trim_sqm) > 0.0001:
                folder_stats_text += f" (Al vivo: {total_trim_sqm:.2f} m²)"

            folder_id = f"folder-{folder_idx}"
            source_html += f'''<div class="folder-container" data-folder-id="{folder_id}">
                <div class="folder-header" style="border-left-color: {current_folder_color};">
                    <span class="folder-stats">{folder_stats_text}</span>
                    <span class="folder-path" style="background-color: {current_folder_color}; color: white; padding: 2px 8px; border-radius: 5px;">{html.escape(display_folder)}</span>
                </div>
                <textarea class="annotation-area folder-annotation" id="anno-{folder_id}" placeholder="Annotazione cartella..."></textarea>'''
            
            subfolder_stats = defaultdict(lambda: {'files': set(), 'pages': 0, 'sqm': 0, 'trim_sqm': 0})
            for page_data in pages:
                subfolder = self._get_display_path(page_data['file_info'])
                if subfolder == ".": continue
                stats = subfolder_stats[subfolder]
                stats['files'].add(os.path.join(page_data['file_info']['path'], page_data['file_info']['filename']))
                stats['pages'] += 1
                page_details = page_data['file_info']['pages_details'][page_data['page_num']]
                stats['sqm'] += page_details.get('area_sqm', 0)
                stats['trim_sqm'] += page_details.get('trim_area_sqm', page_details.get('area_sqm', 0))
        return '<html></html>'

if __name__ == '__main__':
    app = MockFileScannerApp()
    mock_pages = [
        {
            'file_info': {
                'path': 'c:\\test',
                'filename': 'file1.pdf',
                'scan_root': 'c:\\test',
                'pages_details': [
                    {'area_sqm': 1.0, 'trim_area_sqm': 0.9},
                    {'area_sqm': 1.1, 'trim_area_sqm': 1.0}
                ]
            },
            'page_num': 0
        }
    ]
    app._generate_html_content(mock_pages)