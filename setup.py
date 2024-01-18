from setuptools import setup

setup(name='docx-mergefields',
      version='0.1.1',
      description='Replaces fields in docx (Microsoft Office Word) files',
      long_description=open('README.rst').read(),
      classifiers=[
          'License :: OSI Approved :: MIT License',
          'Programming Language :: Python :: 3.7',
          'Topic :: Text Processing',
      ],
      author='Iulian Ciorăscu',
      author_email='ciulian@gmail.com',
      url='http://github.com/iulica/docx-mergefields',
      license='MIT',
      py_modules=['mergefields'],
      zip_safe=False,
      install_requires=['python-docx']
)
