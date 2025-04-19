# Compression Summary Report

Sheet Name | Original Size | Module 1 | Module 2 | Module 3 | Combined | Best Approach
--- | --- | --- | --- | --- | --- | ---
DCF | 5734 tokens | 1.00x | 1.34x | 1.87x | 3.54x | Combined (3.54x)
Income Statement - FCF | 29178 tokens | 1.02x | 1.41x | 1.94x | 4.80x | Combined (4.80x)
Balance Sheet | 8090 tokens | 1.00x | 1.12x | 1.65x | 2.07x | Combined (2.07x)
Variance Analysis | 3584 tokens | 1.00x | 1.14x | 1.85x | 2.38x | Combined (2.38x)
Revenue Build (Toggle & Link) | 38078 tokens | 1.00x | 1.24x | 1.66x | 2.45x | Combined (2.45x)
Revevnue Build Base | 40068 tokens | 1.03x | 1.26x | 1.69x | 2.59x | Combined (2.59x)
Public Comparables | 5655 tokens | 1.00x | 1.13x | 1.89x | 2.41x | Combined (2.41x)
Pre and Post Earnings | 265 tokens | 1.00x | 1.16x | 1.69x | 2.21x | Combined (2.21x)
Forex Assumptions | 15785 tokens | 1.00x | 1.08x | 1.58x | 1.79x | Combined (1.79x)
Stock Price History | 120598 tokens | 1.00x | 1.65x | 1.73x | 5.51x | Combined (5.51x)
Customer Demographics | 1773 tokens | 1.00x | 1.22x | 2.53x | 4.61x | Combined (4.61x)

## Recommendations

Based on the test results, here are the recommended approaches:

1. **For financial sheets with numeric data**: Use Module 3 (Format-Aware Aggregation)
2. **For text-heavy sheets**: Use Module 2 (Inverted Index)
3. **For very large sheets**: Use chunked encoding with Module 3
4. **For small sheets with varied formats**: Use the Combined approach

## Implementation Details

- **Module 1**: Identifies and preserves important rows/columns while discarding uniform data
- **Module 2**: Groups identical values to reduce token usage for repeated data
- **Module 3**: Groups cells with similar formats for better compression
- **Combined**: Applies all three modules in sequence

Large sheets are automatically chunked with headers preserved in each chunk.