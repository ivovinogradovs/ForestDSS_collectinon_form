# ForestDSS Collection Form

A standalone browser-based form for characterizing Forest Decision Support Systems (DSS)
according to the ForestDSS ontology.

Developed for the Training School – Forest Ecosystem Decision Support Services for Climate
Resilience and Decarbonisation, organized jointly by COST Action CA22141 (DSS4ES) and OPTFOR-EU.

## Usage

Open the form directly in any modern browser – no installation required:

**[Launch form](https://ivovinogradovs.github.io/ForestDSS_collectinon_form/forestdss_form.html)**

Or download `forestdss_form.html` and open it locally.

## What it does

The form guides users through characterizing a DSS instance by instance across all
ForestDSS ontology classes:

- Decision Support System
- Software
- Model (vegetation dynamics, natural disturbances, habitat, carbon, wildlife, NWFP, soil)
- Forest Management Strategy
- Ecosystem Services (full CICES hierarchy with cascading dropdowns)
- Planning Problem
- Case Study
- Input and Output Data
- Decision Support Techniques (MADM/MODM)
- Uncertainty Evaluation
- Scenario
- Participatory Process
- Lessons Learned
- Knowledge Management Processes
- Actor / Person

All controlled vocabularies follow the ForestDSS ontology and the working definitions
established in the DSS semantics working document (DSS_semantics_2025).

## Export

Completed characterizations can be exported as:
- **JSON** (recommended) – full nested structure, suitable for wiki import
- **Flat CSV** – one row per field, all sections combined, suitable for spreadsheet analysis

## Ontology reference

Based on the ForestDSS domain ontology (Reuter et al., in review), implemented in OWL and
evaluated using the HermiT reasoner. The ontology comprises 59 classes, 36 object properties,
and 87 datatype properties.

## Repository structure

    forestdss_form.html    standalone form application
    README.md              this file
