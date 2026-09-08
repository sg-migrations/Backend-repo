package github.workflow

import rego.v1

deny contains msg if {
    not input.on.pull_request
    msg := "Workflow must run on pull_request for PR validation."
}
