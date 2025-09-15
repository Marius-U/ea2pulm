# ea2pulm

The `ea2puml.py` script is a command-line tool that exports diagrams from Enterprise Architect (EA) to PlantUML format (.puml). Here’s a summary of its functionality:

## Purpose

- **Export the currently selected EA diagram** (from the EA Project Browser) to a PlantUML .puml file.
- **Preserves layout**, colors, and structure as much as possible, including packages, elements, connectors, and notes.

##Key Features

- **Element Grouping:** Groups elements inside package blocks based on their geometry (center-in-rect or area-overlap ≥ 50%).
- **Color Preservation:** Optionally preserves EA element and connector colors in the PlantUML output.
- **Lollipop Interfaces:** Renders interfaces as lollipops by default (configurable).
- **Block Notes:** Exports notes and tagged values as PlantUML notes, preserving newlines.
- **Alias Modes:** Supports different identifier modes for PlantUML references (uuid, human, or name).
- **Edge Labels:** Configurable connector labels (name, stereotype, both, or none).
- **Ordering:** Packages are ordered top-to-bottom, then left-to-right, to match EA’s visual layout.
- **Robust Endpoint Resolution:** Handles connector endpoints even if elements are hidden or missing.
- **Output Directory:** Default output is output, but this is configurable.

## How It Works

1. **Connects to EA via COM:** Uses `pywin32` to connect to a running EA instance.  
2. **Finds the Selected Diagram:** Requires the user to select a diagram in EA before running.  
3. **Gathers Elements and Geometry:** Collects all diagram objects, their types, names, colors, positions, and relationships.  
4. **Groups Elements into Packages:** Based on geometry, determines which elements are inside which packages.  
5. **Collects Connectors:** Gathers all visible connectors, resolves their endpoints, and collects their properties.  
6. **Renders PlantUML:** Outputs PlantUML syntax for packages, elements, notes, and connectors, preserving layout and style as much as possible.  
7. **Writes Output:** Saves the `.puml` file to the specified directory.

## CLI Options

- `-o, --outdir`: Output directory.  
- `-f, --filename`: Output file name (without extension).  
- `-t, --include-tags`: Include tagged values as notes.  
- `-c, --no-colors`: Disable color output.  
- `-s, --element-stereo`: How to display stereotypes.  
- `--edge-labels`: Connector label policy.  
- `-d, --direction`: Layout direction (LR/TB).  
- `-e, --explore`: Print diagnostic info.  
- `--skin`: Include PlantUML skinparams.  
- `--autolayout`: Add PlantUML autolayout directive.  
- `--alias-mode`: How to generate PlantUML identifiers.  
- `--interface-style`: Interface rendering style.  

---

## Typical Usage

1. Open EA, select a diagram.  
2. Run the script from the command line:  

	```bash
	python ea2puml.py -o output_folder
	```
	
	```bash
	python ea2puml.py --skin --direction LR --include-tags --alias-mode human
	```
3. The script connects to EA, exports the selected diagram as a .puml file, and saves it in the specified folder.

## In summary:

This script is a robust, configurable exporter from Enterprise Architect diagrams to PlantUML, preserving much of the visual and structural information from EA.
