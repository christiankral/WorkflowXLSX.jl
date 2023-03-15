# WorkflowXLSX.jl

Read and write physical quantities from XLSX files

## Installation

The following installation of the package has to be performed only once:

```julia
using Pkg
Pkg.add(url = "https://github.com/christiankral/WorkflowXLSX.jl")
```

## Example template

The file `template/Ohm.jl` includes the application of Ohm's law with voltage and resistance read from the XLSX file `template/Ohm.xlsx`

- Place the two file `template/Ohm.jl` and `template/Ohm.xlsx` in your working directory, e.g. `/work`.
- Start Julia
- Install `WorkflowXLSX` if not yet done (see Installation)
- Switch to your working directory, e.g. `cd("/work")`
- Run the test file by `include("Ohm.jl")`
- Your working directory should now include the file `Ohm_result.xlsx`; this file now includes the calculation result of I1

## Note

- Make sure your Julia variables do not interfere with unit names introduced by `Unitfuly.DefaultSymbols`, e.g., do not use `A` for the area of cross section as `A` represents the unit *Ampere*
