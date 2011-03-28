from setuptools import setup, find_packages
from os.path import join

name = 'xlsx_filler'

readme = open('README.txt').read()
history = open('HISTORY.txt').read()

setup(name = name,
      version = 0.1,
      description = 'Tool for filing xslx files with data, without removing '\
          'anything',
      long_description = readme[readme.find('\n\n'):] + '\n' + history,
      keywords = 'Excel xlsx xlsm',
      author = 'Patrick Gerken',
      author_email = 'gerken@patrick-gerken.de',
      url = '',
      download_url = '',
      license = 'GPL version 2',
      packages = '',
      include_package_data = True,
      platforms = 'Any',
      zip_safe = False,
      install_requires=[
        'setuptools',
        'lxml',
      ],
      test_suite="tests",
      tests_require=['unittest2'],
      classifiers = [
      ],
      entry_points = '''
      ''',
)
