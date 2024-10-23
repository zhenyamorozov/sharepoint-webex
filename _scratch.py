import pkg_resources

target_package = 'msal'

# Get a list of all installed distributions
dists = [dist for dist in pkg_resources.working_set]

# Loop over each distribution and check if the target package is a requirement
for dist in dists:
    for req in dist.requires():
        if req.key == target_package:
            print(f'{dist.project_name} requires {target_package}')