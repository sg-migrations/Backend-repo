package main

import rego.v1

deny contains msg if {
    input.permissions.contents == "write"
    msg := "Global permissions: write-all is not allowed."
}
