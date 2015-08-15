# mailstop-sort
A tool for producing usefully-formatted lists from TSV (tilde-separated-values) input.


__What is the key purpose of this project?__

To produce and maintain a tool that is useful for those providing office services for an important firm, known here only as "Client X".


__Will Client X's identity be revealed?  Will this project contain any of Client X's data?__

Absolutely not.

Not only will none of Client X's data be revealed, certain pieces of code will be kept out of the repository because they would reveal details of the formatting of Client X's data.  This precaution is probably unnecessary, but we will take it anyhow to make sure we stay on the right side of things.


__Why share this project, then?__

The problems being solved on behalf of Client X aren't unique; others may face the same problems, and this tool may be useful for them as well.

With proper encapsulation, we should be able to limit the pieces that deal with the specifics of Client X's data, and thus can't be held in the repository, to just one function.  Anyone else who wants to use the rest of our code to deal with their own tables will simply need to write a function with the same interface (which would probably be needed to deal with the differing format, anyways.)


__What platform and language is the tool?__

The platform will be Windows computers, so the tool is a .wsf (Windows Script File).  Some portions must be written in VBScript in order to have access to the filesystem.  As much of the rest as possible will be written in JScript.
