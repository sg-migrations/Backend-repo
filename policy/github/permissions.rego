package main

import rego.v1


# --------------------------------
# Policy 1 - Global permissions
# --------------------------------

deny contains msg if {
    input.permissions.contents == "write"
    msg := "Global contents: write permission is not allowed."
}


# --------------------------------
# Policy 2 - DEV must follow Build
# --------------------------------

deny contains msg if {
    input.jobs["deploy-dev"].needs != "build"
    msg := "DEV deployment must depend on the build job."
}


# --------------------------------
# Policy 3 - QA must follow DEV
# --------------------------------

deny contains msg if {
    input.jobs["deploy-qa"].needs != "deploy-dev"
    msg := "QA deployment must depend on DEV deployment."
}


# --------------------------------
# Policy 4 - PROD must follow QA
# --------------------------------

deny contains msg if {
    input.jobs["deploy-prod"].needs != "deploy-qa"
    msg := "Production deployment must depend on QA deployment."
}


# --------------------------------
# Policy 5 - Evidence after PROD
# --------------------------------

deny contains msg if {
    input.jobs["generate-evidence"].needs != "deploy-prod"
    msg := "Release evidence must run after production deployment."
}
