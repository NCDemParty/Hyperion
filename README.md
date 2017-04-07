# Hyperion
A set of tools to optimize the counting and tallying of votes at conventions and party meetings

## Goal
To create a product that allows a diverse user base to report, count, and tally votes at various party conventions in a standardized, flexible, and optimized way. 

## Project Timeline
Ongoing, meetings and conventions are a permenant part of NCDP's structure so there will always be a relevant need for tools to assist in the process. Meeting schedules for various conventions are outlined in the [North Carolina Democratic Party Plan of Organization](https://drive.google.com/file/d/0B3VE_bcfX7g1Nlkyd083TXp2aW8/view).

## Tools
- The VBA script requires Excel
- [NCDP's Plan of Organization](https://drive.google.com/file/d/0B3VE_bcfX7g1Nlkyd083TXp2aW8/view) to understand how vote strength is allocated and when meetings occur
- Google Sheets to utilize the Google Scripts function

## Field Definitions
- Precinct: Precinct name supplied by the NCSBE, with a "P " appended to prevent Excel from autoformatting as dates
- Cooper Vote: The total vote Roy Cooper received in that precinct
- County Conv Vote: It's important to read **North Carolina Democratic Party, Plan of Organization** § 5.00, art. 2 (2017) to understand the rules that govern vote strength for County Conventions, **North Carolina Democratic Party, Plan of Organization** § 6.00, art. 1 *Allocation of Votes* (2017) for District Conventions, and **North Carolina Democratic Party, Plan of Organization** § 6.00, art. 2 *Allocation of Votes* (2017) for State Conventions.
- Votes Per Delegate: A calculated field that assigns a raw vote strength per individual, this is determined by how many people are present at the convention for County Conventions. 
- Delegates Present: End users will update this column based off of attendance at the meeting, dynamically changing the Votes Per Delegate column.
- Number of Candidates: This is the total number of candidates running for a particular office. This cell informs the scripts to create a unique row for each candidate.
- Position Being Elected: A validation list of offices to be elected that dynamically names sheets created by the ballot creation tool

## Vote Weights
- Weights will be different for each type of convention (County, District, State), so this formula will have to change depending on the convention it's utilized for. Currently, it grabs the raw vote strength per person from the Votes Per Delegate column on the Check In sheet and multiplies that number by the total votes per candidate within that precinct. 
- In accordance with the [North Carolina Democratic Party Plan of Organization](https://drive.google.com/file/d/0B3VE_bcfX7g1Nlkyd083TXp2aW8/view) it may be required to reduce the total allocated votes granted to a particular precinct/county if the district failed to elect their full voting strength, so it's important to provide the end user the ability to edit these on the fly. 

## Project Contributors
- @JessePresnell
- Matthew Pepper
