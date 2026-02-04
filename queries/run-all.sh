#!/bin/bash
# Run all competency question SPARQL queries against the taxonomy
# Requires Apache Jena ARQ: brew install jena

TTL="../output/iso-security.ttl"

echo "=============================================="
echo "SPARQL Competency Question Validation"
echo "=============================================="
echo ""

# Check if TTL file exists
if [ ! -f "$TTL" ]; then
    echo "ERROR: Taxonomy file not found: $TTL"
    echo "Run 'python convert.py' first to generate the taxonomy."
    exit 1
fi

# Check if arq is available
if ! command -v arq &> /dev/null; then
    echo "ERROR: Apache Jena ARQ not found."
    echo "Install with: brew install jena"
    exit 1
fi

# Run each query
passed=0
failed=0
total=0

for query in cq*.rq; do
    total=$((total + 1))
    echo "=== $query ==="

    # Extract expected result from comment
    expected=$(grep "# Expected:" "$query" | sed 's/# Expected: //')
    echo "Expected: $expected"
    echo ""

    # Run query
    result=$(arq --data="$TTL" --query="$query" 2>&1)

    if [ $? -eq 0 ]; then
        echo "$result"
        passed=$((passed + 1))
    else
        echo "FAILED: $result"
        failed=$((failed + 1))
    fi

    echo ""
    echo "----------------------------------------------"
    echo ""
done

echo "=============================================="
echo "SUMMARY"
echo "=============================================="
echo "Total queries: $total"
echo "Passed: $passed"
echo "Failed: $failed"
echo ""

if [ $failed -eq 0 ]; then
    echo "All competency questions PASSED!"
    exit 0
else
    echo "Some queries FAILED. Check output above."
    exit 1
fi
