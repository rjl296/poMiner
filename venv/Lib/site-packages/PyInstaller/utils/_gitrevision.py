#
# The content of this file will be filled in with meaningful data
# when creating an archive using `git archive` or by downloading an
# archive from github, e.g. from github.com/.../archive/develop.zip
#
rev = "ca01921f19"     # abbreviated commit hash
commit = "ca01921f19a268cf6fce157779b2aba043a1a43c"  # commit hash
date = "2020-06-12 19:26:22 +0200"   # commit date
author = "Hartmut Goebel <h.goebel@crazy-compilers.com>"
ref_names = "HEAD -> develop"  # incl. current branch
commit_message = """depend, building: Allows to override hooks from entry-points.

Hook directories from command line option ``--additional-hooks-dir`` now take
precedence over those from package entry-points. PyInstaller hooks still
have lowest priority. This allows users to override the entry-points (as
he/she could override PyInstaller hooks.

* building: Change order directories are collected.
* depend: Update a comment.
* docs: document it
"""
