GITHUB_OWNER=rspott
GITHUB_TOKEN=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
GITHUB_PROJECT=WAF-test02

GITHUB_OWNER_ESC="\\\"$GITHUB_OWNER\\\""
GITHUB_PROJECT_ESC="\\\"$GITHUB_PROJECT\\\""

#find all the open issues and then delete them.
LOOKUP='query {
  viewer {
    id
  }
  repository(name: $GITHUB_PROJECT_ESC, owner: $GITHUB_OWNER_ESC) {
    issues(filterBy: {states: OPEN}, first: 100) {
      edges {
        node {
          id
        }
      }
    }
  }
}'
LOOKUP="$(echo $LOOKUP)"   # the query should be a one-liner, without newlines
COUNT=1
for ISSUE in $(curl -s -H 'Content-Type: application/json' \
    -H "Authorization: bearer $GITHUB_TOKEN" \
    -X POST -d "{ \"query\": \"$LOOKUP\"}" https://api.github.com/graphql | jq --raw-output '.data.repository.issues.edges[].node.id' | sed 's/= /\n/g')
do
    echo $COUNT
    COUNT=$((COUNT+1))

    ISSUE="\\\"$ISSUE\\\""

    DELETE="mutation {
        deleteIssue(input: {issueId: $ISSUE}) {
            clientMutationId
        }
    }"

    DELETE="$(echo $DELETE)" # the query should be a one-liner, without newlines

    curl -s -H 'Content-Type: application/json' \
    -H "Authorization: bearer $GITHUB_TOKEN" \
    -X POST -d "{ \"query\": \"$DELETE\"}" https://api.github.com/graphql
done

#find all the labels we may have created and delete them
while read LABELS; do
  LABELS=${LABELS// /%20}
  echo "$LABELS"
  curl -i -u $GITHUB_OWNER:$GITHUB_TOKEN -X DELETE "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_PROJECT/labels/$LABELS"

done <labels.txt

# Delete all the milestones
for MILESTONE in {1..30}
do
  echo $MILESTONE
   curl -i -u $GITHUB_OWNER:$GITHUB_TOKEN -X DELETE "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_PROJECT/milestones/$MILESTONE"
done


