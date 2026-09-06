#!/usr/bin/env python3
"""
Update HTML files: replace .jpg/.png → .webp in image references.
Skips og:image, twitter:image, favicon.png, apple-touch-icon.png,
og-image.jpg, og-image-original.jpg, walbrugge_logo_round.png, walbrugge_logo_round.webp
"""

import os
import re
import glob

HTML_GLOB = 'public/**/*.html'

KEEP_EXT = {
    'favicon.png', 'apple-touch-icon.png',
    'og-image.jpg', 'og-image-original.jpg',
    'walbrugge_logo_round.png', 'walbrugge_logo_round.webp',
}

def is_keep_line(line):
    lower = line.lower()
    return 'og:image' in lower or 'twitter:image' in lower

def is_keep_filename(fname):
    return fname in KEEP_EXT

def update_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original = content
    lines = content.split('\n')
    new_lines = []
    changed = False

    for line in lines:
        if is_keep_line(line):
            new_lines.append(line)
            continue

        new_line = line

        # Pattern 1: <img src="/path/file.jpg">  →  <img src="/path/file.webp">
        def repl_img_src(m):
            pre = m.group(1)  # src="/path/file
            ext = m.group(2).lower()
            fname = os.path.basename(pre) + '.' + ext
            if is_keep_filename(fname):
                return m.group(0)
            return f'{pre}.webp"'

        new_line = re.sub(
            r'(src="/[^"]*)\.(jpg|jpeg|png)"',
            repl_img_src,
            new_line,
            flags=re.IGNORECASE
        )

        # Pattern 2: background-image:url('/path/file.jpg')  →  url('/path/file.webp')
        def repl_bg_url(m):
            prefix = m.group(1)  # background-image:url('
            path = m.group(2)    # /assets/img/file
            ext = m.group(3).lower()
            suffix = m.group(4)  # ')
            fname = os.path.basename(path) + '.' + ext
            if is_keep_filename(fname):
                return m.group(0)
            return f'{prefix}{path}.webp{suffix}'

        new_line = re.sub(
            r'(background-image:\s*url\([\'"])([^\'"]*)\.(jpg|jpeg|png)([\'"]\))',
            repl_bg_url,
            new_line,
            flags=re.IGNORECASE
        )

        # Pattern 3: <source srcset="/path/file.jpg">  →  srcset="/path/file.webp"
        def repl_srcset(m):
            pre = m.group(1)  # srcset="/path/file
            ext = m.group(2).lower()
            fname = os.path.basename(pre) + '.' + ext
            if is_keep_filename(fname):
                return m.group(0)
            return f'srcset="{pre}.webp"'

        new_line = re.sub(
            r'(srcset="/[^"]*)\.(jpg|jpeg|png)"',
            repl_srcset,
            new_line,
            flags=re.IGNORECASE
        )

        if new_line != line:
            changed = True
        new_lines.append(new_line)

    if changed:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(new_lines))
        return True
    return False


def main():
    html_files = sorted(glob.glob(HTML_GLOB, recursive=True))
    updated = []

    for fp in html_files:
        if not os.path.isfile(fp):
            continue
        if update_file(fp):
            updated.append(fp)
            print(f"  Updated: {fp}")

    print(f"\nHTML bestanden bijgewerkt: {len(updated)}")


if __name__ == '__main__':
    main()
