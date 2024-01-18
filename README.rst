=================
docx Merge Fields
=================

.. image:: https://badge.fury.io/py/docx-mergefields.png
    :alt: PyPI
    :target: https://pypi.python.org/pypi/docx-mergefields

Udpates the MAILMERGE fields in Office Open XML (docx) files. Can be used on any
system without having to install Microsoft Office Word. Supports Python 3.7 and up.
For the moment only the INCLUDEPICTURE fields are supported.

It supports local images, URLs and base64 image encoded strings. Also, it allows 
the resize of the image both in width and height.

The fields are *replaced* so afterwards they cannot be updated anymore.

This library makes use of the excellent `python-docx`_ library.

Also, it is better used after mailmerging the INCLUDEPICTURE fields using the 
`docx-mailmerge2`_ library.


Installation
============

Installation with ``pip``:
::

    $ pip install docx-mergefields


Usage
=====

Open the file.
::

    from mergefields import MergeFieldsDocument
    with MergeFieldsDocument("../documents/mailmerge_doc.docx") as doc:
      doc.transform_fields()
      doc.doc.save("../documents/merged_doc.docx")


Examples
========

From local file, original size:
::

  { INCLUDEPICTURE "./filename.jpg" }

From a URL, resize width to the given size, resize height to 
preserving the aspect ratio:
::

  { INCLUDEPICTURE "https://www.pngall.com/wp-content/uploads/8/Sample-Watermark-PNG-Image.png" \w 200 }

From a base64 URI, resize both width and height to the given value
(aspect ratio is not preserved):
::

  { INCLUDEPICTURE "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAQMAAADCCAMAAAB6zFdcAAAAtFBMVEX///9QUU/yX1xKS0mKiopNTkxCQ0FGR0VpamjS0tJHSEa9vbyvr67yXFlSU1FDRELyWVaoqKjc3N3xU1D19fXDxMNWV1XxUU797Ovv7+/j4+P3pKP5vLv+8/O3t7b71NP1i4nzc3D4rqz1hYPzZ2WCgoGUlJN3eHf0e3lgYV/83t2enp76xsXMzcyQkY/1k5H6ysn2k5L3qqg3ODX2m5rzb21lZmT5t7Xr39/su7rmqana4+QStYMYAAAEt0lEQVR4nO3ab3eiOBQGcGkkWFoEBUVtpbXV+qdqZ9S2uzvf/3stuQkQQNvdOT2j1Of3Yk9NMJhLckkyW6sBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICwWN7eLhfDA7VBuxNF0ar/R3/Sl3ObzaZ9daDy7t4Jfc/zQ2+23FPdmTLuWpbFefd5UKycG3HLTSPIFTbiQnb9BT/7S3HTMNz9MViMQu9CcXynGIVVlzNDMW0+yfe2zami0HQjvpt1ejGIf6i1NwYPoXOhCx9z1XURPQ2zL/XqtYyP2ch9qWIxeAmTIZCEwhtr1Q1LjQDGVCxMrgUhsFVoeG6SVCsGO1/22/fGr/F/5IfXtHoi+2i780l9aqohofX3yhVREcPjSW+1UjH4EcossFuIT733C4qC91NVR1xmgZZ8I1xOqcvbrL+is+abCIKpN1upGNAo8Ea9tOCFSsI7+iBHOnvL3onX3LC6q/TjQDTrDsRlbkdrtkox2IjH7oz0IpocqqglkoE516uj7bP26TnuPZvW6nFiNKdaeZViIEd+L1c2ErkxpLnRFYOct3PVudzXNGkAXIrWubaCqlAM7sQz9zb7Cm9qaqSzyeFGOzQVajJYtjZAKhSDGzEO/F6hdBwPBEe8Gijpu6vaQVNTvRBa8Zwwm1lFhWIwE70dFy/dUZIYytluuEGxPtXnyVShEcOzaFUoBvTEH4uXviejQ6Q6wzrcJj19mTFFp1mWFSsUA5H9nPvipUuREML/EINu1tNITJttOmQqFIMD48BX42D9cQzk20D1my5Ne12hGHyQD7yhHOu5V17eRKwK5leRcEULxjQrVigGn7wXOm55V5wJ5Drakmgnke6mKhSDH7QUeMgXLqhwF//Vd8u74sy1ZRSxuqqrUAzkdsHJn6DNnHTDILpSHAjpp7lZikGaHaoUA5oM3kwv2tB+Qe6eaR1o2Ppieb1VB0l0gGTaGRESK5JXVSkGQ7lvnGUj4UbfN8pnbbJ0SxhMXIM1adavKSNetxLX4tpkf3W6MWDPlzqx+7ml8wPPeaDEOFy+yqDcq2+p00Le6MRvh2DQYozODyLVYK6btLJWGyxaMj2V7nZk1BnGNduWKH9UR0f+eDYbe748T/PSr7VkEEzXEodo8lCJdYOkx/p7M7Cy4yRKJPm75c6ZjoOX8pdND3GYHSmnB4rJTBBa21LyFyGoveUXx4JYLxiM/myU8qW+qzyWQzFQp2k6ei+mOtnJOo0ITh2Xm6RO7ha0bpRvkdOMwZYV8ZasefHyISguHIM6t5IjZca7st9PPG7BLNzDtBmzaDnxZpfudgIxmNSLJslTvCj8+8Ki+N0gmpqWy127u17pzRUTf0TNihzxVL7boeXmabjLzQZ/s/ei/mAwOHySUH07bTY4r59f/y2Ns9lQ2kGdi0U6G/z3Y/+Wo9n4aibMPr/22xo55z0ThB7NhvD22L/jqMQholc6WjwzM4fOEM9X0O//5Yd/9/vfeRn0kX5b+PXP4y/64wQ2+n9c0C6o+P+B93uC/iDp/2BwlhEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAID/61/qyEpXq0NhcQAAAABJRU5ErkJggg==" \w 200 \h 200 }


Todo / Wish List
================

* Include other fields
* Update the fields instead of replacing them

Contributing
============

* Fork the repository on GitHub and start hacking
* Create / fix the unit tests
* Send a pull request with your changes

Unit tests
----------

In order to make sure that the library performs the way it was designed, unit
tests are used. When providing new features, or fixing bugs, there should be a
unit test that demonstrates it. Run the test suite::

    python -m unittest discover

Credits
=======

| This repository is written and maintained by `Iulian Ciorăscu`_.

.. _python-docx: https://pypi.org/project/python-docx/
.. _docx-mailmerge2: https://github.com/iulica/docx-mailmerge
.. _Iulian Ciorăscu: https://github.com/iulica/
