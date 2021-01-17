$TagsTable = @{}

$TagsTable['a'] = [array] "test"


$tag_key = 'a'

$tag_existing_values = $TagsTable[$tag_key]
$tag_new_values = $tag_existing_values + 'edward'
$TagsTable[$tag_key] = $tag_new_values

$TagsTable

$TagsTable['b'] = [array] " another "

$TagsTable
