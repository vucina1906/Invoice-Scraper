from setuptools import setup,find_packages


def get_requirements(file_path:str)->List[str]:
    requirements=[]
    with open(file_path) as file_obj:
        for line in file_obj:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if line.startswith('-e'):
                continue
            requirements.append(line)
    return requirements
            
            

setup(
    name='Invoice Scraper',
    version='1.0.0.',
    author='VukDjunisijevic',
    author_email='vucina19931906@gmail.com',
    packages=find_packages(),
    install_requires=get_requirements('requirements.txt')     
)