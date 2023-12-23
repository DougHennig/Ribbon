# How to contribute to Ribbon

## Report a bug
- Please check [issues](https://github.com/DougHennig/Ribbon/issues) if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement
- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.
- Make whatever changes are necessary.
- Use [FoxBin2Prg](https://github.com/fdbozzo/foxbin2prg) to create text files for all VFP binary files (SCX, VCX, DBF, etc.)
- Update ChangeLog.md and describe the changes. Also, make any necessary changes to the documentation in README.md.
- If you haven't already done so, install VFPX Deployment: choose Check for Updates from the Thor menu, turn on the checkbox for VFPX Deployment, and click Install.
- Run the VFPX Deployment tool to create the installation files: choose VFPX Project Deployment from the Thor Tools, Application menu. Alternately, execute EXECSCRIPT(_screen.cThorDispatcher, 'Thor_Tool_DeployVFPXProject').
- Commit the changes.
- Push to your fork.
- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: 2023-12-23