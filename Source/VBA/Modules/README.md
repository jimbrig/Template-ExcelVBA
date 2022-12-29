# Standard VBA Modules - `.bas` Files

> This directory houses all exported `.bas` Standard VBA Module files for the VBA Project. Standard modules should be named using a `mod*` prefix (i.e. `modUtils.bas`)

## Overview

By default most projects should include, at least the following *base* standard modules:

1. *Globals* module (`modGlobals.bas`) to store global, project-wide constants, enumerations, variables, definitions, functions, metadata, etc.
2. Module for Testing (`modTests.bas`)
3. Module for a Default *Error Handler* (`modErrorHandler.bas`)
4. Module for Exporting VBA (`modExportVBA.bas`)
5. Module for Formatting (`modFormats.bas`)
6. **Build** Module (`modBuild.bas`)
7. **Debug** Module (`modDebug.bas`)
8. Module for Logging Mechanism/Framework (`modLogger.bas` or `modLogging.bas` or `modLogs.bas`, etc.)
9. Optimization Module (`modOptimize.bas`)
10. etc.

## Modules

- [modGlobals.bas](modGlobals.bas): Global, Project-Wide Declarations, Enumerations, Functions, Metadata, etc.
- [modTests.bas](modTests.bas): Testing module
- [modErrorHandler.bas](modErrorHandler.bas): Error Handling Module
- [modExportVBA.bas](modExportVBA.bas): Development Module to Export VBA

- ...

