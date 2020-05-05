####################################################################################################
# Copyright (C) University of Tübingen, Germany - all rights reserved.
# Unauthorized copying of this file, via any medium is strictly prohibited.
# Proprietary and confidential.
# Authors: Leon Kuchenbecker <kuchenb@informatik.uni-tuebingen.de>
#          Leon Bichmann <bichmann@informatik.uni-tuebingen.de>
####################################################################################################

from setuptools import setup, find_packages

requires = [
        'pandas',
        'xlrd',
        ]


setup(
        name = 'assignmenttool',
        version = '1.0.0',
        package_dir={'':'src'},
        packages=find_packages('./src'),
        author = 'Leon Kuchenbecker',
        author_email = 'leon.kuchenbecker@uni-tuebingen.de',
        description = 'Assignment Tool',
        install_requires = requires,
        zip_safe = False,
        include_package_data=True,
        entry_points={
            'console_scripts' : [
                'assignment-tool          = assignmenttool:main',
                ],
            },
        )
