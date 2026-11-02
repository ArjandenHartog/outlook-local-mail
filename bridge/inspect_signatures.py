import base64
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
PATH = Path(os.environ['APPDATA']) / 'Microsoft/Signatures'


def read_file(path):
    raw = path.read_bytes()
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16')
    if raw.startswith(b'\xef\xbb\xbf'):
        return raw.decode('utf-8-sig')
    try:
        return raw.decode('utf-8')
    except UnicodeDecodeError:
        return raw.decode('cp1252')


class Images(HTMLParser):
    def __init__(self):
        super().__init__()
        self.images = []
    def handle_starttag(self, tag, attrs):
        if tag.lower() in ('img', 'v:imagedata'):
            self.images.append(dict(attrs))


def main():
    output = ROOT / 'runtime/signature-images'
    output.mkdir(exist_ok=True)
    rows = []
    for path in PATH.glob('*.htm'):
        parser = Images()
        html = read_file(path)
        parser.feed(html)
        row = {'name':path.stem,'images':[],'html_length':len(html),'text':read_file(path.with_suffix('.txt')) if path.with_suffix('.txt').exists() else ''}
        for index, image in enumerate(parser.images):
            src = image.get('src','')
            detail = {'alt':image.get('alt'),'src_start':src[:70],'width':image.get('width'),'height':image.get('height')}
            match = re.fullmatch(r'data:(image/[^;]+);base64,(.+)',src,re.S)
            if match:
                suffix = {'image/png':'.png','image/jpeg':'.jpg','image/gif':'.gif'}.get(match[1],'.bin')
                image_path = output / (path.stem + '-' + str(index) + suffix)
                image_path.write_bytes(base64.b64decode(match[2]))
                detail['file'] = str(image_path)
            row['images'].append(detail)
        rows.append(row)
    (ROOT/'runtime/signature-inventory.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(rows,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    main()
