Associate Office365 Data
========================

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](http://www.apache.org/licenses/LICENSE-2.0) ![This script was last tested in Nuix 8.6](https://img.shields.io/badge/Script%20Tested%20in%20Nuix-8.6-green.svg)

View the GitHub project [here](https://github.com/Nuix/Associate-Office365-Data) or download the latest release [here](https://github.com/Nuix/Associate-Office365-Data/releases).

# Overview

This script helps associate `manifest.xml` **File** data entries to the corresponding Nuix item's in a case created from ingesting the corresponding O365 data.

**IMPORTANT NOTE:** This script currently only supports file data and not email!  This is because while each file entry has a **SHA-256**, I have not determined an identifier I can use to associate email entries back to email items in Nuix.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

# Usage

After ingesting some O365 data into Nuix with **SHA-256** digest generation enabled, run this script.  Select the relevant `manifest.xml` file and select whether you would like custodian values applied, whether you want updated items to have tags applied and the name of the tag to apply.  The script will use the **SHA-256** values for **File** documents in the XML to apply the **Tag** values in the XML file as custom metadata on the matched items.

# Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the script.

# License

```
Copyright 2020 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
