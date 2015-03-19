(inactive project, utomatically exported from code.google.com/p/yet-another-fake-basic)


Yet Another Fake Basic compiler, which tries to be compatible with VB6's syntax and ABI (Application Binary Interface). This is an experimental project, and may or may not be compatible with all features of VB6, and may or may not be development further.

## Technical Information

The compiler is written in VB6, using LLVM (Low Level Virtual Machine Compiler Infrastructure, http://www.llvm.org/) library for back end and code generation support. The lexer is manually constructed in a single function, which looks like a DFA but it isn't. The parser is manually constructed Recursive-Descent Analyzer. The AST (Abstract Syntax Tree) is object-oriented constructed, all nodes are implemented from a base class. The symbol table is implemented using Collection in VB6.
