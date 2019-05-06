# Contributing to CivilConnection and CivilPython

You are welcome to contribute and help make the toolset better and more complete! 
The project follows the [fork & pull](https://help.github.com/articles/using-pull-requests/#fork--pull) model for accepting contributions.
As a contributor, here are the guidelines we would like you to follow:

# Contributor Code of Conduct

As contributors and maintainers of the project, we pledge to respect everyone who contributes by posting issues, updating documentation, submitting pull requests, providing feedback in comments, and any other activities.

Communication through any of channels must be constructive and never resort to personal attacks, trolling, public or private harassment, insults, or other unprofessional conduct.


# Filing issues

When filing an issue, make sure to follow the following steps:

1. Which version are you using?
2. What did you do? (Steps to reproduce)
3. What did you expect to happen?
4. What did happen instead?
5. If helpful, include a screenshot. Annotate the screenshot for clarity.

## Branching Strategy 

CivilConnection supports multiple releases of Revit and Civil 3D starting from 2017. There exists one main branch for each release to allow bug fixes for previous releases. However main development happens on the master branch.
In general, contributors should develop on branches based off of master and pull requests should be made against master. 

**Fork and clone the repository.**
* Create a new branch based on master: 
    `git checkout -b <my-branch-name> master`
* Make your changes, test them and make nothing breaks.
* Push to your fork and submit a pull request from your branch to master.
* Pat yourself on the back and wait for your pull request to be reviewed.

## Submitting a pull request


Here are a few things you can do that will increase the likelihood of your pull request to be accepted:

* Follow the existing style where possible.
* Keep your change as focused as possible. 
* If you want to make multiple independent changes, please consider submitting them as separate pull requests.
* Write a good commit message.

## Commit Messages

Please include a brief description of the change you made.  If it is based off an issue then mention this reference at the beginning of the commit message.

Also do your best to factor commits appropriately, not too large with unrelated things in the same commit, and not too small with the same small change applied N times in N different commits.


