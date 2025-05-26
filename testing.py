from corankco import BordaCount, ScoringScheme, Dataset

# Each ranking is a list of candidates from most to least preferred
rankings = [
    ['A', 'B', 'C', 'D'],  # Voter 1
    ['B', 'A', 'D', 'C'],  # Voter 2
    ['C', 'D', 'B', 'A']   # Voter 3
]

# Create a BordaCount aggregator
borda = BordaCount()
dataset = Dataset(rankings)
scoring_scheme = ScoringScheme([[0., 1., 1., 0., 1., 0.], [1., 1., 0., 1., 1., 0.]])
# Aggregate the rankings
result = borda.compute_consensus_rankings(dataset, scoring_scheme)

# Print the result
print("Borda Count Result:")
for rank, candidate in enumerate(result, start=1):
    print(f"{rank}. {candidate}")