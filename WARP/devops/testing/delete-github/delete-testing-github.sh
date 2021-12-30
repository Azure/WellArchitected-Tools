GITHUB_OWNER=rspott
GITHUB_TOKEN=$1
GITHUB_PROJECT=WAF-test02

GITHUB_OWNER_ESC="\\\"$GITHUB_OWNER\\\""
GITHUB_PROJECT_ESC="\\\"$GITHUB_PROJECT\\\""

DATA='{"name":"'
DATA+=$GITHUB_PROJECT
DATA+='"}'

# Delete the entire repo
curl \
  -i -X DELETE \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: token ${GITHUB_TOKEN}" \
   https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_PROJECT}

sleep 5

# create a new one!
curl \
  -X POST \
  -H "Accept: application/vnd.github.v3+json" \
  -H "Authorization: token ${GITHUB_TOKEN}" \
  https://api.github.com/user/repos \
  -d "$DATA"

