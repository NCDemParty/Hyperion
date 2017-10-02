drop table if exists cooper_county;

CREATE LOCAL TEMPORARY TABLE cooper_county
ON COMMIT PRESERVE ROWS AS /*+DIRECT*/ 
select county, SUM(case when contest_name='NC GOVERNOR' and candidate_party='DEM' then total_votes else null end) as cooper_votes 
from election_reporting.precinct_20161108
group by 1
order by 1;

drop table if exists dalton_county;

CREATE LOCAL TEMPORARY TABLE dalton_county
ON COMMIT PRESERVE ROWS AS /*+DIRECT*/ 
select county, SUM(case when contest_name='NC GOVERNOR' and candidate_party='DEM' then total_votes else null end) as dalton_votes 
from election_reporting.precinct_20121106
group by 1
order by 1;

drop table if exists cooper_cd;

CREATE LOCAL TEMPORARY TABLE cooper_cd
ON COMMIT PRESERVE ROWS AS /*+DIRECT*/ 
select UPPER(county_name) as county, ARRAY_AGG(CONCAT(cd, cd_votes)) as cd_votes FROM
(select f.county_name, CONCAT(CONCAT('CD ',ush_district), ': ') as cd, ROUND(SUM(gov_dem_vote)::float/300::float)::float as cd_votes from legion.full_results f
group by 1, 2) s
group by 1
order by 1;


select c.county, ROUND(cooper_votes/300,0) as state_convention,
cd_votes,
case when ROUND((dalton_votes::float/3000::float)::float)<2 then 2 else ROUND((dalton_votes::float/3000::float)::float) END as current_sec,
case when ROUND((cooper_votes::float/2309128)*450::float)+1<2 then 2 else ROUND((cooper_votes::float/2309128)*450::float)+1 end as new_sec
from cooper_county c
left join dalton_county d on c.county=d.county
left join cooper_cd cd on cd.county=d.county and cd.county=c.county;


select UPPER(f.county_name), prec_id, enr_desc, case when ROUND(SUM(gov_dem_vote)::float/100::float)::float<1 then 1 else ROUND(SUM(gov_dem_vote)::float/100::float)::float END as county_convention_votes
from legion.full_results f
where prec_id is not NULL
group by 1, 2,3;
