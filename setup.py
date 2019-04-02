from setuptools import setup

setup(name='robotframework-ms-outlook',
      version='0.0.1',
      description='Keywords dealing with MS Outlook',
      author='DPA',
      packages=['MsOutlook'],
      install_requires=['robotframework==3.1.1', 'extract-msg==0.22.1'],
      zip_safe=False)
