# SW-Export-BOMs
Little badly writen macro that I created so that you can export all the BOMs of a single assembly

I got fed up with clicking "Save As" for every single configuration in an assembly, so I put this together. It’s not fancy, and the code might give a professional developer a headache, but it gets your Parts Only BOMs into Excel without the manual soul-crush.

##What it actually does
- **Mass Export:** Can do just the active config or blast out every single configuration in the file at once.
- **Custom Naming:** You can tell it which **Custom Properties** to use for the filename **_(so you don't have to rename 50 files later)._**
- **Parts Only:** It skips the sub-assembly headers and just gives you the raw parts list.

##How to use it
- Pick your BOM template **_.sldbomtbt_**
- Select where you want the files to go.
- Type the names of the properties you want in the filename **eg., _PartNumber_, _Revision_**
- **Run it**

##Just a heads up
**References:** If the folder browser crashes, make sure you have "Microsoft Shell Controls And Automation" checked in your VBA References.

**Excel:** You need Excel installed (obviously).

**Template:** If your template is broken, the export will be broken.

##Contributing
If you want to make the code less "badly written," feel free to tweak it.