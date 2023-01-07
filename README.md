# Excel VBA Project Template

## Project Structure

The directory structure and file extension utilized in this template are opinionated and not necessarilly the best for
all projects. 

The goal of this template is to provide a starting point for a project that is easy to understand and maintain, while 
remaining comprehensive in terms of supporting a variety of project components and adhering to best practices. 

The primary root level folders are:

- [Build](./Build) - Contains the built binary workbooks and their corresponding zip archives for each incremental build
    performed. Builds are versioned using [Semantic Versioning]() and fall under project-specific types such as 
    *Debug*, *Test*, *Dev*, *Prod*, etc. The subfolders of the build directory represent the versions and types of that
    build (i.e. `Build/v1.0.1/Debug/DBG_MyProject.xlsm`, `Build/v1.0.1/Test/TEST_MyProject.xlsm`, etc.).

    Ensure that the build directory is added to the `.gitignore` file to prevent the build artifacts from being committed.
    The build artifacts should be included as release assets in the hosted, version controlled (GitHub) repository.

- [Documentation](./Documentation) - Contains the project documentation. This includes all relevant docs and assets
    pertaining to the the project allocated into sub-folders respectively. I have included common folders such as 
    [Admin](./Documentation/Admin/) for Administrative documents (i.e. emails, attachments, proposals, contracts, etc.),
    [Help](./Documentation/Help/) if the project contains a Microsoft HTML Help (`.chm`) document, 
    [Site](Documentation/Site/) for a dedicated, hosted website (currently supports `mkdocs`), etc. 

- [Development]()

- [Source]()

### Directory Skeleton

<details><summary>Click to View Project Directory Tree</summary><p>

```powershell

```

</p></details>
