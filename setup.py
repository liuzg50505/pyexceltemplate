from setuptools import setup, find_packages

setup(
    name="wysiwygtemplate",
    version="1.0",
    keywords=("wysiwyg", "excel", "template", "excel template"),
    description="A wysiwyg excel template library!",
    long_description="A wysiwyg excel template library!",
    license="MIT Licence",

    url="http://test.com",
    author="liuzg",
    author_email="liuzg50505@qq.com",

    packages=find_packages(),
    include_package_data=True,
    platforms="any",
    install_requires=['openpyxl'],

    scripts=[],
    entry_points={
        'console_scripts': [
            'test = test.help:main'
        ]
    }
)