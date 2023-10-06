# Excel VBA Project Template

> ![NOTE] A template project for Visual Basic for Applications (VBA) in Excel.

<!-- Start:Badges -->

[![License](https://img.shields.io/badge/License-Unlicense-yellow?logo=license)](https://unlicense.org/)
[![Version](https://img.shields.io/badge/Version-0.0.1-blue?logo=version)](https://semver.org/)

<!-- End:Badges -->

## Overview

This is a template for a project usings [Visual Basic for Applications (VBA)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel),
specifically in [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel).

Use Cases:

- [ ] Project Setup | Initialization | Configuration
- [ ] Project Build | Compilation
- [ ] Project Deployment | Distribution
- [ ] Project Testing | Validation
- [ ] Project Maintenance | Support
- [ ] Project Documentation | User Guide | Tutorial

## Requirements

- [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)
- [Microsoft Visual Studio Code](https://code.visualstudio.com/)
- [Microsoft Visual Basic for Applications (VBA) Language Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Microsoft Visual Basic for Applications (VBA) Object Library Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/library-reference)

## Features

- [ ] Project Setup | Initialization | Configuration
- [ ] Project Build | Compilation
- [ ] Project Deployment | Distribution
- [ ] Project Testing | Validation
- [ ] Project Maintenance | Support
- [ ] Project Documentation | User Guide | Tutorial

## Usage

1. Create new repository from this template.

2. Clone the repository to your local machine and setup git's remote origin and branching strategy.

```bash
git clone <REPOSITORY_URL>
git remote add origin <REMOTE_ORIGIN_URL>
git flow init
git flow config set push.default simple
git pull origin main
```

3. Open the project in Visual Studio Code.

```bash
code <PROJECT_PATH>
```

4. Review core files:

- [config.yml](config.yml): Contains project configuration information such as project name, author, description, etc.

- [deploy.txt](deploy.txt): Contains the path to an existing directory indicating the location where the production version of the application is to be deployed/distributed to.

- [README.md](README.md): The project's README file.

- [LICENSE](LICENSE.md): The project's license file.

- [CHANGELOG.md](CHANGELOG.md): The project's change log file.



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
