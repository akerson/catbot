import openpyxl
from catbotUnit import Unit
import discord
import datetime
import asyncio
from discord.ext.commands import Bot
from discord.ext import commands
from itertools import islice
import math

updatedate = "October 13th"
clientID = "foobar" #grab your discord bot id and put it here

Client = discord.Client()
bot_prefix= "!"
client = commands.Bot(command_prefix=bot_prefix)
unitlist = {}

def chunks(l, n):
    # For item i in a range that is a length of l,
    for i in range(0, len(l), n):
        # Create an index range for l of n items:
        yield l[i:i+n]

def generateblock(full,title):
	#take a list of 3 lists and generate a codeblock back
	cols = 3
	s = "```ini\n"
	s += '{:=^65}\n'.format(title)
	rows = len(full[0]) #full[0] because it has to be the longest
	for row in range(rows):
		for col in range(cols):
			if len(full[col]) > row: #aka we have something to print
				s += "{:22}".format(full[col][row])
		s += "\n" # nenrow before next one
	s += "```"
	return s

@client.event
async def on_ready():
	print("Bot Online!")
	print("Name: {}".format(client.user.name))
	print("ID: {}".format(client.user.id))

@client.command(pass_context=True)
async def lookup(ctx, *, name: str):
	print("T: {0} | S: {1} | C: #{2} | A: {3} | M: {4}".format(datetime.datetime.now(), ctx.message.server.name, ctx.message.channel.name, ctx.message.author.name, name))
	slvl = ""
	lvl = 0
	if name[0] == "[" and "]" in name:
		#we have a level present! save it and set name equal to the remainder
		slvl = name[name.find("[")+1:name.find("]")]
		lvl = int(slvl)
		name = name[name.find("]")+1:]
		if name[0] == " ":
			name = name[1:]
	name = name.lower()
	if name is None or name not in unitlist:
		await client.say("I'm sorry, I don't recognize that unit. Unit list available in !lookuphelp")
		print("unit {0} not in list.".format(name))
	else:
		try:
			unit = unitlist[name]
			embed=discord.Embed(title="Type", description=unit.getType(), color=unit.getRarityColor())
			embed.set_author(name=unit.getName(lvl),icon_url=unit.getIcon())
			embed.set_thumbnail(url=unit.getPic())
			embed.add_field(name="Unlocks From", value=unit.getUnlocksFrom(), inline=True)
			embed.add_field(name="AP - Upkeep", value=unit.getAPUpkeep(), inline=True)
			embed.add_field(name="Move - Run Speed", value=unit.getMoveRunSpd(), inline=True)
			embed.add_field(name="HP + HP/Lvl", value=unit.getHP(lvl), inline=True)
			embed.add_field(name="Damage 1", value=unit.getDMG1(lvl), inline=True)
			embed.add_field(name="Priority 1", value=unit.getPriority(1), inline=True)
			embed.add_field(name="Damage 2", value=unit.getDMG2(lvl), inline=True)
			embed.add_field(name="Priority 2", value=unit.getPriority(2), inline=True)
			embed.add_field(name="Targets", value=unit.getTargets(), inline=True)
			embed.add_field(name="Traits", value=unit.getTraits(), inline=False)
			embed.add_field(name="Resists", value=unit.getResists(), inline=True)
			embed.set_footer(text="Hi I'm a catbot. Type '!lookuphelp' for information on how to use me! [Updated: {}]".format(updatedate))
			await client.say(embed=embed)
		except:
			print("ERROR! Please investigate!!")
			await client.say("I'm sorry, something went wrong (likely a typo on our end)! Error logged, we'll look into it!")
			raise

@client.command()
async def lookuphelp():
	rnames = []
	dnames = []
	enames = []
	for key in unitlist:
		if unitlist[key].faction == "Republic":
			rnames.append(unitlist[key].name)
		elif unitlist[key].faction == "Dominion":
			dnames.append(unitlist[key].name)
		else:
			enames.append(unitlist[key].name)
	rnames.sort()
	dnames.sort()
	enames.sort()
	rn = int(math.ceil(len(rnames)/3))
	dn = int(math.ceil(len(dnames)/3))
	en = int(math.ceil(len(dnames)/3))
	rlonglist = list(chunks(rnames,rn))
	dlonglist = list(chunks(dnames,dn))
	elonglist = list(chunks(enames,en))

	output = "is here to help!\n\n**How to Use**\n!lookup *unitname* (eg: !lookup soldier)\n!lookup *[lvl]unitname* (eg: !lookup [10]a.p.c.)\n\n**Current Units in Game:**\n"
	output += generateblock(rlonglist,"[REPUBLIC]")
	await client.say(output)
	output = generateblock(dlonglist,"[DOMINION]")
	await client.say(output)
	output = generateblock(elonglist,"[EMPIRE]")
	await client.say(output)

wb = openpyxl.load_workbook('data.xlsx')
sht = wb.get_sheet_by_name('Sheet1')

#make all of them
print("Populating Database...")
for row in range(2, sht.max_row + 1):
	newunit = Unit()
	newunit.name = sht['A' + str(row)].value
	newunit.utype = sht['B' + str(row)].value
	newunit.rarity = sht['C' + str(row)].value
	newunit.unlockedbyunit = sht['D' + str(row)].value
	newunit.unlockedbylvl = sht['E' + str(row)].value
	newunit.maxlvl = sht['F' + str(row)].value
	newunit.squad = sht['G' + str(row)].value
	newunit.apcost = sht['H' + str(row)].value
	newunit.upkeep = sht['I' + str(row)].value
	newunit.movespd = sht['J' + str(row)].value
	newunit.runspd = sht['K' + str(row)].value
	newunit.deploytime = sht['L' + str(row)].value
	newunit.hitsinf = sht['M' + str(row)].value
	newunit.hitshi = sht['N' + str(row)].value
	newunit.hitsvehicle = sht['O' + str(row)].value
	newunit.hitstank = sht['P' + str(row)].value
	newunit.hitsheli = sht['Q' + str(row)].value
	newunit.hitsplane = sht['R' + str(row)].value
	newunit.hitsbase = sht['S' + str(row)].value
	#newunit.BLANK = sht['T' + str(row)].value
	newunit.basehp = sht['U' + str(row)].value
	newunit.hplvl = sht['V' + str(row)].value
	#newunit.maxhp = sht['W' + str(row)].value
	#newunit.BLANK = sht['X' + str(row)].value
	newunit.d1type = sht['Y' + str(row)].value
	newunit.d1dmg = sht['Z' + str(row)].value
	newunit.d1dmglvl = sht['AA' + str(row)].value
	newunit.d1range = sht['AB' + str(row)].value
	newunit.d1spread = sht['AC' + str(row)].value
	newunit.d1clip = sht['AD' + str(row)].value
	newunit.d1aim = sht['AE' + str(row)].value
	newunit.d1fire = sht['AF' + str(row)].value
	newunit.d1reload = sht['AG' + str(row)].value
	#newunit.cooldown = sht['AH' + str(row)].value
	#newunit.d1dps = sht['AI' + str(row)].value
	newunit.d1hitsinf = sht['AJ' + str(row)].value
	newunit.d1hitshi = sht['AK' + str(row)].value
	newunit.d1hitsvehicle = sht['AL' + str(row)].value
	newunit.d1hitstank = sht['AM' + str(row)].value
	newunit.d1hitsheli = sht['AN' + str(row)].value
	newunit.d1hitsplane = sht['AO' + str(row)].value
	newunit.d1hitsbase = sht['AP' + str(row)].value
	#newunit.BLANK = sht['AQ' + str(row)].value
	newunit.d1hitsinfP = sht['AR' + str(row)].value
	newunit.d1hitshiP = sht['AS' + str(row)].value
	newunit.d1hitsvehicleP = sht['AT' + str(row)].value
	newunit.d1hitstankP = sht['AU' + str(row)].value
	newunit.d1hitsheliP = sht['AV' + str(row)].value
	newunit.d1hitsplaneP = sht['AW' + str(row)].value
	newunit.d1hitsbaseP = sht['AX' + str(row)].value
	#newunit.BLANK = sht['AY' + str(row)].value
	newunit.d2type = sht['AZ' + str(row)].value
	newunit.d2dmg = sht['BA' + str(row)].value
	newunit.d2dmglvl = sht['BB' + str(row)].value
	newunit.d2range = sht['BC' + str(row)].value
	newunit.d2spread = sht['BD' + str(row)].value
	newunit.d2clip = sht['BE' + str(row)].value
	newunit.d2aim = sht['BF' + str(row)].value
	newunit.d2fire = sht['BG' + str(row)].value
	newunit.d2reload = sht['BH' + str(row)].value
	#newunit.cooldown = sht['BI' + str(row)].value
	#newunit.d1dps = sht['BJ' + str(row)].value
	newunit.d2hitsinf = sht['BK' + str(row)].value
	newunit.d2hitshi = sht['BL' + str(row)].value
	newunit.d2hitsvehicle = sht['BM' + str(row)].value
	newunit.d2hitstank = sht['BN' + str(row)].value
	newunit.d2hitsheli = sht['BO' + str(row)].value
	newunit.d2hitsplane = sht['BP' + str(row)].value
	newunit.d2hitsbase = sht['BQ' + str(row)].value
	#newunit.BLANK = sht['BR' + str(row)].value
	newunit.d2hitsinfP = sht['BS' + str(row)].value
	newunit.d2hitshiP = sht['BT' + str(row)].value
	newunit.d2hitsvehicleP = sht['BU' + str(row)].value
	newunit.d2hitstankP = sht['BV' + str(row)].value
	newunit.d2hitsheliP = sht['BW' + str(row)].value
	newunit.d2hitsplaneP = sht['BX' + str(row)].value
	newunit.d2hitsbaseP = sht['BY' + str(row)].value
	#newunit.BLANK = sht['BZ' + str(row)].value
	newunit.trait1 = sht['CA' + str(row)].value
	newunit.trait2 = sht['CB' + str(row)].value
	newunit.trait3 = sht['CC' + str(row)].value
	#newunit.BLANK = sht['CD' + str(row)].value
	newunit.resistsinf = sht['CE' + str(row)].value
	newunit.resistshi = sht['CF' + str(row)].value
	newunit.resistsvehicle = sht['CG' + str(row)].value
	newunit.resiststank = sht['CH' + str(row)].value
	newunit.resistsplane = sht['CI' + str(row)].value
	newunit.resistsheli = sht['CJ' + str(row)].value
	newunit.resistsbase = sht['CK' + str(row)].value
	newunit.faction = sht['CL' + str(row)].value
	unitlist[newunit.name.lower()] = newunit
print("Database complete.")

client.run(clientID)
