import tbapy
import json

key = 'Kzpr14G0WK8hAz1Eoa7Tk3gybFAjutmOxJUEuNIwPbmDfW75Eqvkgp9YLN9MgGis'
tba = tbapy.TBA(key)


team = tba.team(230)
districts = tba.team_districts(1418)
#tba.district_rankings(district)
#match = tba.match(year=2017, event='chcmp', type='sf', number=2, round=1)
events = tba.team_events(230, 2019)
robots = tba.team_robots(230)

#print(json.dumps(team, indent=4, sort_keys=True))
#print(json.dumps(districts, indent=4, sort_keys=True))
print(json.dumps(events, indent=4, sort_keys=True))
print(json.dumps(robots, indent=4, sort_keys=True))

print "done"
