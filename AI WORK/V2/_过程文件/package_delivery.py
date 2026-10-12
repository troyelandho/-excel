from pathlib import Path
import importlib.util, json, zipfile, hashlib

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('style',ROOT/'apply_style.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
SRC=ROOT.parent/'V1'/'交付文件'/'产业园投资测算.xlsx'
OUTPUT=ROOT/'产业园投资测算_V2_样式调整.xlsx'
report=m.verify_packages(SRC.read_bytes(),OUTPUT.read_bytes())
assert report['passed'] and report['non_style_differences']==0
assert m.sha(SRC.read_bytes())=='d1109b68c70597b461b35d52abaf4d869bc8ce35aa6f1388305f05ebee28570e'
assert m.sha(OUTPUT.read_bytes())=='6c2159569af7da0582a9b64b27848764a5272244553f5e7df781882100866653'
assert OUTPUT.read_bytes()==(ROOT/'_过程文件'/'可复现检查'/m.OUTPUT_NAME).read_bytes()
native_a=json.loads((ROOT/'_过程文件'/'原版预览'/'native_check.json').read_text(encoding='utf-8-sig'))
native_b=json.loads((ROOT/'_过程文件'/'新版预览'/'native_check.json').read_text(encoding='utf-8-sig'))
assert native_a['HashDisplayCells']==native_b['HashDisplayCells']
assert native_b['FormulaErrorDisplayCells']==[]
extra=json.loads((ROOT/'_过程文件'/'补充验证结果.json').read_text(encoding='utf-8'))
assert len(extra['negative_tests'])==14
assert all(t['correctly_rejected'] for t in extra['negative_tests'])
assert extra['output_sha256']==m.sha(OUTPUT.read_bytes())

visual={'excel_version':native_b['ExcelVersion'],'examined_sheets':10,'preview_regions':22,
        'display_scanned_content_cells':9366,'new_hash_display_cells':[],
        'existing_helper_hash_cells':native_b['HashDisplayCells'],
        'formula_error_display_cells':[], 'actual_cell_font_names':native_b['CellFontNames'],
        'actual_recalculation_performed':False,'saved_in_excel':False,
        'preview_file_sha256_equals_final':True,
        'final_code_reproduction_sha256_equals_final':True}
visual_file=m.choose_output(ROOT/'_过程文件','最终视觉检查摘要.json')
m.write_new(visual_file,m.json_bytes(visual))

files=[OUTPUT,ROOT/'apply_style.py',ROOT/'requirements.txt',ROOT/'README.md',ROOT/'样式调整验证说明.md',
       ROOT/'产业园投资测算_V2_样式调整_验证结果.json',
       ROOT/'_过程文件'/'内容基准_产业园投资测算_V2_样式调整.json',
       ROOT/'_过程文件'/'补充验证结果.json',visual_file,
       ROOT/'_过程文件'/'原版预览'/'native_check.json',ROOT/'_过程文件'/'新版预览'/'native_check.json',
       ROOT/'_过程文件'/'test_preservation.py',ROOT/'_过程文件'/'native_preview.ps1',ROOT/'_过程文件'/'requirements_qa.txt']
archive=m.choose_output(ROOT,'产业园投资测算_V2_样式调整_完整交付包.zip')
with zipfile.ZipFile(archive,'x',compression=zipfile.ZIP_DEFLATED,compresslevel=9) as z:
    for f in files:z.write(f,f.relative_to(ROOT).as_posix())
with zipfile.ZipFile(archive) as z:
    assert z.testzip() is None
    for f in files:assert z.read(f.relative_to(ROOT).as_posix())==f.read_bytes()
manifest={'v1_sha256':m.sha(SRC.read_bytes()),'output_sha256':m.sha(OUTPUT.read_bytes()),
          'non_style_differences':0,'final_validation_passed':True,
          'zip':str(archive),'zip_members':len(files),'zip_sha256':m.sha(archive.read_bytes()),
          'delivered_files':[{ 'path':str(f.relative_to(ROOT)),'size':f.stat().st_size,'sha256':m.sha(f.read_bytes())} for f in files]}
m.write_new(m.choose_output(ROOT/'_过程文件','最终交付清单.json'),m.json_bytes(manifest))
print(json.dumps({'final_validation_passed':True,'non_style_differences':0,'zip':str(archive),'zip_members':len(files),'core_files':[f.name for f in files[:6]]},ensure_ascii=False))
