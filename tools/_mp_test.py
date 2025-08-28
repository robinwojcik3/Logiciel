import sys, os
QGIS_ROOT = r"C:\Program Files\QGIS 3.40.3"
qgis_py_root = os.path.join(QGIS_ROOT, 'apps', 'Python312')
qgis_py_lib = os.path.join(qgis_py_root, 'Lib')
qgis_site   = os.path.join(qgis_py_lib, 'site-packages')
qgis_app_py = os.path.join(QGIS_ROOT, 'apps', 'qgis', 'python')
old_syspath = list(sys.path)
sys.path = [qgis_py_root, qgis_py_lib, qgis_site, qgis_app_py] + [p for p in old_syspath if isinstance(p,str) and '.venv' not in p.lower()]
print('sys.path head (old):', sys.path[:5])
try:
    import concurrent.futures.process
    import multiprocessing.connection
    import _multiprocessing
    print('Old: _multiprocessing OK')
except Exception as e:
    print('Old: _multiprocessing FAILED:', type(e).__name__, e)
qgis_dlls = os.path.join(qgis_py_root, 'DLLs')
sys.path = [qgis_py_root, qgis_py_lib, qgis_dlls, qgis_site, qgis_app_py] + [p for p in old_syspath if isinstance(p,str) and '.venv' not in p.lower()]
print('sys.path head (new):', sys.path[:6])
try:
    import importlib
    if '_multiprocessing' in sys.modules:
        del sys.modules['_multiprocessing']
    _multiprocessing = importlib.import_module('_multiprocessing')
    print('New: _multiprocessing OK')
except Exception as e:
    print('New: _multiprocessing FAILED:', type(e).__name__, e)
