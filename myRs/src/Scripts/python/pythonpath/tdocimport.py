#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# Created by modifying urlimport.py in 10.11.loading_modules_from_a_remote_machine_using_import_hooks of Python Cookbook 3rd Edition.
# 一旦LibreOfficeを終了させないとimportはキャッシュが使われるのでデバッグ時は必ずLibreOfficeを終了すること!!!
# インポートするパッケジーには__init__.pyが必要。
import sys
import importlib.abc
from types import ModuleType
def _get_links(simplefileaccess, url):
	foldercontents = simplefileaccess.getFolderContents(url, True)  # フルパスで返ってくる。
	tdocpath = "".join((url, "/"))
	return [content.replace(tdocpath, "") for content in foldercontents]
class UrlMetaFinder(importlib.abc.MetaPathFinder):  # meta path finderの実装。
	def __init__(self, simplefileaccess, baseurl):
		self._simplefileaccess = simplefileaccess
		self._baseurl = baseurl
		self._links   = {}
		self._loaders = {baseurl: UrlModuleLoader(simplefileaccess, baseurl)}
	def find_module(self, fullname, path=None):  # find_moduleはPython3.4で撤廃だが、find_spec()にしてもそのままではうまく動かない。
		if path is None:
			baseurl = self._baseurl
		else:
			if not path[0].startswith(self._baseurl):
				return None
			baseurl = path[0]
		parts = fullname.split('.')
		basename = parts[-1]
		if basename not in self._links:  # Check link cache
			self._links[baseurl] = _get_links(self._simplefileaccess, baseurl)
		if basename in self._links[baseurl]:  # Check if it's a package。パッケージの時。
			fullurl = "/".join((self._baseurl, basename))
			loader = UrlPackageLoader(self._simplefileaccess, fullurl)
			try:  # Attempt to load the package (which accesses __init__.py)
				loader.load_module(fullname)
				self._links[fullurl] = _get_links(self._simplefileaccess, fullurl)
				self._loaders[fullurl] = UrlModuleLoader(self._simplefileaccess, fullurl)
			except ImportError:
				loader = None
			return loader
		filename = "".join((basename, '.py'))
		if filename in self._links[baseurl]:  # A normal module
			return self._loaders[baseurl]
		else:
			return None
	def invalidate_caches(self):
		self._links.clear()
class UrlModuleLoader(importlib.abc.SourceLoader):  # Module Loader for a URL
	def __init__(self, simplefileaccess, baseurl):
		self._simplefileaccess = simplefileaccess
		self._baseurl = baseurl
		self._source_cache = {}
	def module_repr(self, module):
		return '<urlmodule {} from {}>'.format(module.__name__, module.__file__)
	def load_module(self, fullname):  # Required method
		code = self.get_code(fullname)
		mod = sys.modules.setdefault(fullname, ModuleType(fullname))
		mod.__file__ = self.get_filename(fullname)
		mod.__loader__ = self
		mod.__package__ = fullname.rpartition('.')[0]
		exec(code, mod.__dict__)
		return mod
	def get_code(self, fullname):  # Optional extensions
		src = self.get_source(fullname)
		return compile(src, self.get_filename(fullname), 'exec')
	def get_data(self, path):
		pass
	def get_filename(self, fullname):
		return "".join((self._baseurl, '/', fullname.split('.')[-1], '.py'))
	def get_source(self, fullname):
		filename = self.get_filename(fullname)
		if filename in self._source_cache:
			return self._source_cache[filename]
		try:
			inputstream = self._simplefileaccess.openFileRead(filename)
			dummy, b = inputstream.readBytes([], inputstream.available())  # simplefileaccess.getSize(module_tdocurl)は0が返る。
			source = bytes(b).decode("utf-8")  # モジュールのソースをテキストで取得。
			self._source_cache[filename] = source
			return source
		except:
			raise ImportError("Can't load {}".format(filename))
	def is_package(self, fullname):
		return False
class UrlPackageLoader(UrlModuleLoader):  # Package loader for a URL
	def load_module(self, fullname):
		mod = super().load_module(fullname)
		mod.__path__ = [self._baseurl]
		mod.__package__ = fullname
	def get_filename(self, fullname):  # パッケージの時はまず__init__.pyを実行。
		return "/".join((self._baseurl, '__init__.py'))
	def is_package(self, fullname):
		return True
_installed_meta_cache = {}  # meta path finderを入れておくグローバル辞書。重複を防ぐ目的。
def install_meta(simplefileaccess, address):  # Utility functions for installing the loader
	if address not in _installed_meta_cache:  # グローバル辞書にないパスの時。
		finder = UrlMetaFinder(simplefileaccess, address)  # meta path finder。モジュールを探すクラスをインスタンス化。
		_installed_meta_cache[address] = finder  # グローバル辞書にmeta path finderを登録。
		sys.meta_path.append(finder)  # meta path finderをsys.meta_pathに登録。
def remove_meta(address):  # Utility functions for uninstalling the loader
	if address in _installed_meta_cache:
		finder = _installed_meta_cache.pop(address)
		sys.meta_path.remove(finder)
