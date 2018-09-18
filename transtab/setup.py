import setuptools

long_description = ''

with open('README.md', 'r') as fh:
	long_description = fh.read()


setuptools.setup(
	name='transtab',
	version='0.0.1',
	author='Jithu Sunny',
	author_email='jithusunnyk@gmail.com',
	description='Transform tabular data',
	long_description=long_description,
	long_description_content_type='text/markdown',
	url='https://github.com/SolveForTech/csvprogram',
    packages=setuptools.find_packages(),
    entry_points = {
    	'console_scripts': ['transtab=transtab.transtab_cmd:main'],
    },
    classifiers=(
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    )
)
