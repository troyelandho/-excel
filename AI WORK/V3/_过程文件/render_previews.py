from pathlib import Path
import subprocess,sys
from PIL import Image,ImageOps,ImageDraw
root=Path(__file__).parent/'视觉检查'
exe=Path('C:/Users/BBEN/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe')
thumbs=[]
for p in sorted(root.glob('*.pdf')):
 if '--final-only' not in sys.argv or p.stem[:2] in ['01','02','13']:
  subprocess.run([str(exe),'-f','1','-singlefile','-scale-to','1900','-png',str(p),str(p.with_suffix(''))],check=True,capture_output=True)
 im=Image.open(p.with_suffix('.png')).convert('RGB');im.thumbnail((580,440))
 tile=Image.new('RGB',(600,470),'#eeeeee');tile.paste(im,((600-im.width)//2,25));ImageDraw.Draw(tile).text((8,5),p.stem[:2],fill='black');thumbs.append(tile)
contact=Image.new('RGB',(600*4,470*((len(thumbs)+3)//4)),'white')
for i,t in enumerate(thumbs):contact.paste(t,((i%4)*600,(i//4)*470))
contact.save(root/'全部预览.png')
print('Rendered',len(thumbs),'native Excel previews')
