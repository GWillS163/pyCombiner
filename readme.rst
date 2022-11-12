The function of pyCombiner
=======================

Combine that all your python files in your project sequential into one by the relationship of import statement.


|image0|


How it deals with import statement
==================================

Let me summarize introduce of it simply.

- recognize "import ..." sentence and "from ... import ..." sentence ...
- turn the import sentence into function codes
- combine the function codes into one file

In recent version, it can deal with "import ..." sentence basically, but not enough.

Supported Sentence
==================

.. code:: python
    import os
    import os, sys
    from bs4 import BeautifulSoup  # will completely import even though that function not use exterior lib
    from bs4 import *
    # some libraries are not written by user. so it's line of codes will be preserved.


Not Supported Yet
=================

.. code:: python
   from src import showString as s  # Can't process the "as ..." statements
    from src import showRests  # partially import but use other function


.. |image0| image:: https://github.com/GWillS163/pyCombiner/raw/master/res/introImg.png
