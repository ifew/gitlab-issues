# Use
``npm install``
``node gitlab_sla_analysis.js``

# For cross check
## search project
``curl --header "PRIVATE-TOKEN: xxxTOKENxx" "http://git.domain.com/api/v4/groups?search=project"``

## get projects in group
``curl --header "PRIVATE-TOKEN: xxxTOKENxx" "http://git.domain.com/api/v4/groups/10/projects"``

## get project info
``curl --header "PRIVATE-TOKEN: xxxTOKENxx" "http://git.domain.com/api/v4/projects/56"``

## get project list
``curl --header "PRIVATE-TOKEN: xxxTOKENxx" "http://git.domain.com/api/v4/groups/10/projects?per_page=100"``

## get group issues list
``curl --header "PRIVATE-TOKEN: xxxTOKENxx" "http://git.domain.com/api/v4/groups/10/issues?per_page=100"``