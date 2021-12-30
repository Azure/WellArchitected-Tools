GITHUB_OWNER=rspott
GITHUB_TOKEN=$1
GITHUB_PROJECT=WAF-test02
GITHUB_REPO_ID="MDEwOlJlcG9zaXRvcnk0MDQ4MDQ1NzQ="

GITHUB_OWNER_ESC="\\\"$GITHUB_OWNER\\\""
GITHUB_PROJECT_ESC="\\\"$GITHUB_PROJECT\\\""
GITHUB_REPO_ID_ESC="\\\"$GITHUB_REPO_ID\\\""

COUNT=1
while read ISSUE; do
    echo $COUNT
    COUNT=$((COUNT+1))

    ISSUE="\\\"$ISSUE\\\""
    BODY="\\\"foo\\\""

    NEW="mutation {
  createIssue(input: {repositoryId: $GITHUB_REPO_ID_ESC, title: $ISSUE, body: $ISSUE}) {
    issue {
      number
      body
    }
  }
}"

    NEW="$(echo $NEW)" # the query should be a one-liner, without newlines
# echo "New :   " $NEW

    curl -s -H 'Content-Type: application/json' \
    -H "Authorization: bearer $GITHUB_TOKEN" \
    -X POST -d "{ \"query\": \"$NEW\"}" https://api.github.com/graphql
# exit
sleep 1
done <./testing/delete-github/issues.txt
