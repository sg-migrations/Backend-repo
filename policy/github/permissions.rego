package github.workflow

import rego.v1

deny contains msg if {
    input.permissions == "write-all"
    msg := "Global permissions: write-all is not allowed."
}
