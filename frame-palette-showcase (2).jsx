import { useState, useMemo } from "react";

const DARK = [
  { num:"01", name:'Core Black', tag:'Pure · Minimal · Essential', best:'Developer tools, IDEs', bg:"#121211", surface:"#1C1C1A", elevated:"#272725", text:"#F5F4EF", muted:"#96938E", accent:"#F5C831", accentFg:"#000000", secondary:"#E0E0E0", border:"#33332F" },
  { num:"02", name:'Carbon', tag:'Warm Depth · Refined Darkness', best:'Editorial, docs, premium tools', bg:"#141411", surface:"#201E19", elevated:"#2B2A22", text:"#F2EFE9", muted:"#8E8678", accent:"#F5C831", accentFg:"#000000", secondary:"#C8A428", border:"#37352B" },
  { num:"03", name:'Slate Circuit', tag:'Cool Tech · Digital Precision', best:'Infrastructure, API dashboards', bg:"#121415", surface:"#1C1E20", elevated:"#25282E", text:"#F0EFEE", muted:"#7C818A", accent:"#F5C831", accentFg:"#000000", secondary:"#4A8FE0", border:"#2E343D" },
  { num:"04", name:'Midnight Moss', tag:'Organic Tech · Grounded Energy', best:'Sustainability, health platforms', bg:"#121310", surface:"#1C1F19", elevated:"#262A22", text:"#EEEEE6", muted:"#858778", accent:"#F5C831", accentFg:"#000000", secondary:"#78B830", border:"#30372A" },
  { num:"05", name:'Deep Forge', tag:'Industrial Heat · Molten Edge', best:'IoT, robotics, maker tools', bg:"#131211", surface:"#1F1B19", elevated:"#2A2422", text:"#F2EFEB", muted:"#91817D", accent:"#F5C831", accentFg:"#000000", secondary:"#E05830", border:"#372D2A" },
  { num:"06", name:'Obsidian', tag:'Volcanic Glass · Cold Depth', best:'Creative suites, generative AI', bg:"#121214", surface:"#1D1D1C", elevated:"#27272A", text:"#F1F0EE", muted:"#84818A", accent:"#F5C831", accentFg:"#000000", secondary:"#9090E0", border:"#333335" },
  { num:"07", name:'Gunmetal', tag:'Military Neutral · Solid Ground', best:'Admin panels, enterprise tools', bg:"#131414", surface:"#1F2020", elevated:"#2A2C2C", text:"#F0EFEB", muted:"#858986", accent:"#F5C831", accentFg:"#000000", secondary:"#9AA0A0", border:"#363836" },
  { num:"08", name:'Deep Navy', tag:'Command · Authority · Maritime', best:'Fintech, analytics, dashboards', bg:"#121315", surface:"#1B1F24", elevated:"#242A34", text:"#EEEEEE", muted:"#777D8A", accent:"#F5C831", accentFg:"#000000", secondary:"#3070D0", border:"#2C3442" },
  { num:"09", name:'Petroleum', tag:'Industrial Teal · Raw Precision', best:'Data engineering, pipelines', bg:"#121312", surface:"#1B1F1C", elevated:"#242A28", text:"#ECEFEB", muted:"#72817D", accent:"#F5C831", accentFg:"#000000", secondary:"#28A0A0", border:"#2C3733" },
  { num:"10", name:'Void Purple', tag:'Ethereal · Speculative · Edge', best:'Creative tools, AI/ML platforms', bg:"#131213", surface:"#1F1B1C", elevated:"#2A242A", text:"#F1EFED", muted:"#888186", accent:"#F5C831", accentFg:"#000000", secondary:"#A060E0", border:"#352E39" },
  { num:"11", name:'Cinder', tag:'Ash · Smoke · Post-Burn Calm', best:'Documentation, wikis, reading', bg:"#151412", surface:"#211F1B", elevated:"#2D2A24", text:"#F0EEE9", muted:"#928A7E", accent:"#F5C831", accentFg:"#000000", secondary:"#B8A890", border:"#3A352D" },
  { num:"12", name:'Iron Oxide', tag:'Raw Metal · Rust & Grit', best:'Maker communities, hardware', bg:"#15120F", surface:"#211B16", elevated:"#2E241C", text:"#F2EEE8", muted:"#987C6E", accent:"#F5C831", accentFg:"#000000", secondary:"#B85434", border:"#3A2C24" },
  { num:"13", name:'Volcanic', tag:'Magma Core · Eruptive Energy', best:'Gaming, sports, launch dashboards', bg:"#141210", surface:"#201C19", elevated:"#2D2521", text:"#F2EFE8", muted:"#917D70", accent:"#F5C831", accentFg:"#000000", secondary:"#E84020", border:"#392B25" },
  { num:"14", name:'Umbra', tag:'Shadow Warmth · Quiet Power', best:'Luxury brands, premium portfolio', bg:"#15120F", surface:"#211C16", elevated:"#2E251C", text:"#F2EEE6", muted:"#987E6E", accent:"#F5C831", accentFg:"#000000", secondary:"#B07050", border:"#3A2D24" },
  { num:"15", name:'Steel', tag:'Structural Blue · Cold Clarity', best:'DevOps, CI/CD, cloud monitoring', bg:"#111416", surface:"#1B2022", elevated:"#252D31", text:"#EEEEEC", muted:"#76868C", accent:"#F5C831", accentFg:"#000000", secondary:"#5A98B8", border:"#2D3940" },
  { num:"16", name:'Deep Jungle', tag:'Bio-Digital · Dense & Alive', best:'Environmental tech, agriculture', bg:"#111511", surface:"#1A2119", elevated:"#242E23", text:"#ECEFE6", muted:"#728A72", accent:"#F5C831", accentFg:"#000000", secondary:"#38A060", border:"#2B3A2C" },
  { num:"17", name:'Abyss', tag:'Deep Ocean · Compressed Calm', best:'Focus apps, immersive writing', bg:"#121213", surface:"#1B1B1C", elevated:"#24242A", text:"#ECEAEC", muted:"#727081", accent:"#F5C831", accentFg:"#000000", secondary:"#5050C0", border:"#2B2B35" },
  { num:"18", name:'Nightfall', tag:'Dusk Signal · Purple Twilight', best:'Design tools, creative studios', bg:"#121213", surface:"#1D1B1C", elevated:"#28242A", text:"#F0ECED", muted:"#84788A", accent:"#F5C831", accentFg:"#000000", secondary:"#8050C0", border:"#342B39" },
  { num:"19", name:'Charcoal Rose', tag:'Dark Romance · Structured Warmth', best:'Wellness, beauty, social platforms', bg:"#151312", surface:"#201E1C", elevated:"#2C2727", text:"#F3EEEB", muted:"#918181", accent:"#F5C831", accentFg:"#000000", secondary:"#C06080", border:"#382F30" },
  { num:"20", name:'Eclipse', tag:'Total Dark · Zero Distraction', best:'IDEs, deep-work, terminal apps', bg:"#131313", surface:"#1F1F1D", elevated:"#2A2A29", text:"#F0EFEB", muted:"#858482", accent:"#F5C831", accentFg:"#000000", secondary:"#8888A8", border:"#363636" },
  { num:"21", name:'Copper Vein', tag:'Molten Copper · Warm Industrial', best:'Manufacturing, supply chain', bg:"#14120E", surface:"#211C14", elevated:"#2F2619", text:"#F3EEE4", muted:"#987F62", accent:"#F5C831", accentFg:"#000000", secondary:"#C87830", border:"#3C2F20" },
  { num:"22", name:'Dark Amber', tag:'Honeyed Depth · Liquid Gold', best:'Finance, investment platforms', bg:"#13120F", surface:"#1F1D15", elevated:"#2A281C", text:"#F4F0E2", muted:"#968A67", accent:"#F5C831", accentFg:"#000000", secondary:"#D0A020", border:"#383423" },
  { num:"23", name:'Military', tag:'Tactical Green · Operational Precision', best:'Field ops, logistics, defence tools', bg:"#121310", surface:"#1C2018", elevated:"#262B20", text:"#EDEEE2", muted:"#808670", accent:"#F5C831", accentFg:"#000000", secondary:"#708040", border:"#2F3725" },
  { num:"24", name:'Slate Rose', tag:'Hard Edge · Soft Core', best:'HR tech, workplace tools', bg:"#141315", surface:"#1F1D21", elevated:"#2B2730", text:"#F2EDEA", muted:"#8F8592", accent:"#F5C831", accentFg:"#000000", secondary:"#C08098", border:"#37303A" },
  { num:"25", name:'Burgundy', tag:'Dark Wine · Structured Luxury', best:'Legal tech, wealth management', bg:"#141210", surface:"#201B19", elevated:"#2B2321", text:"#F3ECE8", muted:"#917D78", accent:"#F5C831", accentFg:"#000000", secondary:"#A03040", border:"#372925" },
  { num:"26", name:'Electric Cyan', tag:'Signal Blue · Live Data', best:'Network monitoring, telecom', bg:"#111314", surface:"#1A1F20", elevated:"#232A2D", text:"#EAEFEE", muted:"#6E7D81", accent:"#F5C831", accentFg:"#000000", secondary:"#00B8D4", border:"#292F35" },
  { num:"27", name:'Dark Olive', tag:'Field Green · Grounded Authority', best:'Mapping, geospatial, outdoor tech', bg:"#121310", surface:"#1E2016", elevated:"#282B1D", text:"#EEEEE0", muted:"#848670", accent:"#F5C831", accentFg:"#000000", secondary:"#90A030", border:"#323721" },
  { num:"28", name:'Indigo Depth', tag:'Deep Indigo · Cosmic Logic', best:'Research, data science, academia', bg:"#121214", surface:"#1B1B20", elevated:"#24242E", text:"#ECEAEE", muted:"#77748A", accent:"#F5C831", accentFg:"#000000", secondary:"#6060D0", border:"#2B2B3D" },
  { num:"29", name:'Dark Emerald', tag:'Vault Green · Precious Depth', best:'Fintech, insurance, compliance', bg:"#101512", surface:"#19221C", elevated:"#223025", text:"#E8EEE4", muted:"#6C8A78", accent:"#F5C831", accentFg:"#000000", secondary:"#18B878", border:"#283B30" },
  { num:"30", name:'Frost Night', tag:'Cold Intelligence · Polar Precision', best:'Weather tech, climate data', bg:"#121415", surface:"#1C1F21", elevated:"#252A2E", text:"#ECEEEE", muted:"#77818A", accent:"#F5C831", accentFg:"#000000", secondary:"#80B8D8", border:"#2D363D" },
  { num:"31", name:'Mahogany', tag:'Rich Wood · Enduring Craft', best:'Real estate, architecture', bg:"#15110F", surface:"#211A15", elevated:"#2F231A", text:"#F4EEE6", muted:"#987A6E", accent:"#F5C831", accentFg:"#000000", secondary:"#A8583E", border:"#3A2A22" },
  { num:"32", name:'Cobalt', tag:'Deep Blue · Structural Confidence', best:'Aviation, aerospace, engineering', bg:"#121317", surface:"#1B1D24", elevated:"#242832", text:"#EAEAEE", muted:"#727893", accent:"#F5C831", accentFg:"#000000", secondary:"#3858C8", border:"#2B2F42" },
  { num:"33", name:'Sienna', tag:'Terracotta Dark · Earthen Warmth', best:'Food tech, agriculture, artisan', bg:"#141210", surface:"#211D18", elevated:"#2E261F", text:"#F3EEE7", muted:"#967F70", accent:"#F5C831", accentFg:"#000000", secondary:"#C07048", border:"#3A2F26" },
  { num:"34", name:'Dark Violet', tag:'Deep Violet · Speculative Edge', best:'Crypto, blockchain, DeFi, web3', bg:"#131213", surface:"#1F1B1E", elevated:"#2A242C", text:"#F1ECEE", muted:"#887D8E", accent:"#F5C831", accentFg:"#000000", secondary:"#9858C8", border:"#342B39" },
  { num:"35", name:'Sea Glass', tag:'Coastal Dark · Submerged Calm', best:'Marine tech, oceanography', bg:"#111514", surface:"#1A221F", elevated:"#24302B", text:"#E8EEE8", muted:"#6E8A7D", accent:"#F5C831", accentFg:"#000000", secondary:"#30B898", border:"#293A33" },
  { num:"36", name:'Pewter', tag:'Cool Steel · Quiet Infrastructure', best:'Internal tools, enterprise back-office', bg:"#121416", surface:"#1D2023", elevated:"#282D32", text:"#F0EFEC", muted:"#7E8790", accent:"#F5C831", accentFg:"#000000", secondary:"#7898B0", border:"#333A42" },
  { num:"37", name:'Crimson', tag:'Dark Crimson · Urgent Precision', best:'Security ops, incident response', bg:"#141110", surface:"#201918", elevated:"#2D2020", text:"#F4EAE7", muted:"#9A7D7D", accent:"#F5C831", accentFg:"#000000", secondary:"#D02840", border:"#3A2727" },
  { num:"38", name:'Desert Night', tag:'Warm Sand · Night Sky', best:'Travel tech, outdoor navigation', bg:"#15140F", surface:"#222017", elevated:"#302B1F", text:"#F4F0E6", muted:"#9A8E70", accent:"#F5C831", accentFg:"#000000", secondary:"#C0A05C", border:"#3D3525" },
  { num:"39", name:'Amethyst', tag:'Purple Rose · Structured Mystery', best:'Luxury e-commerce, beauty tech', bg:"#141216", surface:"#201B22", elevated:"#2C2630", text:"#F2ECEE", muted:"#918098", accent:"#F5C831", accentFg:"#000000", secondary:"#B86AC8", border:"#372D3D" },
  { num:"40", name:'Copper Teal', tag:'Oxidised Metal · Patina Precision', best:'Materials science, chemistry', bg:"#121312", surface:"#1B1F1C", elevated:"#242C2A", text:"#ECF0EB", muted:"#778681", accent:"#F5C831", accentFg:"#000000", secondary:"#509898", border:"#2B3830" },
  { num:"41", name:'Onyx', tag:'Pure Black · Absolute Focus', best:'Premium dev tools, stealth UIs', bg:"#121210", surface:"#1B1B19", elevated:"#262624", text:"#F4F2EE", muted:"#96938E", accent:"#F5C831", accentFg:"#000000", secondary:"#C8C8C8", border:"#32322E" },
  { num:"42", name:'Mercury', tag:'Liquid Silver · Polished Logic', best:'Financial terminals, trading', bg:"#141414", surface:"#202020", elevated:"#2C2C2B", text:"#F2F0EC", muted:"#898986", accent:"#F5C831", accentFg:"#000000", secondary:"#B8B8B0", border:"#393936" },
  { num:"43", name:'Glacier', tag:'Ice Deep · Compressed Cold', best:'Climate platforms, scientific data', bg:"#121314", surface:"#1B1E1F", elevated:"#24292C", text:"#ECEEED", muted:"#727D86", accent:"#F5C831", accentFg:"#000000", secondary:"#60A0C8", border:"#2B2F35" },
  { num:"44", name:'Rust Plum', tag:'Iron Meets Wine · Industrial Romance', best:'Maker tools, craft e-commerce', bg:"#151210", surface:"#201B19", elevated:"#2C2422", text:"#F3ECE9", muted:"#9A8681", accent:"#F5C831", accentFg:"#000000", secondary:"#A84870", border:"#372B2C" },
  { num:"45", name:'Stone Blue', tag:'Cut Stone · Cold Certainty', best:'Architecture, engineering portals', bg:"#121316", surface:"#1D1F26", elevated:"#292C38", text:"#EEEDEB", muted:"#7E8294", accent:"#F5C831", accentFg:"#000000", secondary:"#6A78C0", border:"#343746" },
  { num:"46", name:'Dusk Gold', tag:'Evening Gold · Sunset Authority', best:'Publishing, media, content platforms', bg:"#141310", surface:"#202017", elevated:"#2C2B22", text:"#F4F0E6", muted:"#9A8E78", accent:"#F5C831", accentFg:"#000000", secondary:"#D0B040", border:"#383423" },
  { num:"47", name:'Deep Plum', tag:'Rich Plum · Velvet Precision', best:'Music platforms, entertainment tech', bg:"#131210", surface:"#201B1C", elevated:"#2B242A", text:"#F2ECEC", muted:"#91868E", accent:"#F5C831", accentFg:"#000000", secondary:"#A05090", border:"#352B35" },
  { num:"48", name:'Arctic Night', tag:'Polar Dark · Blue-White Cold', best:'Research ops, satellite data', bg:"#111519", surface:"#1A2228", elevated:"#24303A", text:"#ECEDEE", muted:"#748292", accent:"#F5C831", accentFg:"#000000", secondary:"#A4BFE8", border:"#2C3A48" },
  { num:"49", name:'Bronze', tag:'Ancient Metal · Enduring Signal', best:'Heritage brands, museum tech', bg:"#14130F", surface:"#211E15", elevated:"#2F281A", text:"#F3EFE6", muted:"#96886D", accent:"#F5C831", accentFg:"#000000", secondary:"#B89040", border:"#3A3222" },
  { num:"50", name:'Dark Lagoon', tag:'Deep Teal · Subterranean Calm', best:'Diving tech, aquaculture, deep-sea data', bg:"#111412", surface:"#1A201C", elevated:"#232D28", text:"#E8EEE9", muted:"#6C867D", accent:"#F5C831", accentFg:"#000000", secondary:"#289888", border:"#273830" },
  { num:"51", name:'Graphite Bloom', tag:'Soft Graphite · Warm Signal', best:'Collaboration suites, team portals', bg:"#141312", surface:"#211F1C", elevated:"#2D2A25", text:"#F2F0E8", muted:"#968C7A", accent:"#F5C831", accentFg:"#000000", secondary:"#B8A45A", border:"#3A352D" },
  { num:"52", name:'Storm Glass', tag:'Rain Blue · Polished Calm', best:'Status pages, operations dashboards', bg:"#121518", surface:"#1B2024", elevated:"#252C32", text:"#EDF1F2", muted:"#7A8790", accent:"#F5C831", accentFg:"#000000", secondary:"#6FA7C8", border:"#2D3840" },
  { num:"53", name:'Ink Fern', tag:'Botanical Dark · Quiet Systems', best:'Knowledge bases, healthcare admin', bg:"#121511", surface:"#1B211A", elevated:"#252D24", text:"#ECF1E8", muted:"#7C8B78", accent:"#F5C831", accentFg:"#000000", secondary:"#6FA85A", border:"#2E3A2C" },
  { num:"54", name:'Lunar Copper', tag:'Moonlit Metal · Warm Precision', best:'Hardware platforms, manufacturing ops', bg:"#141210", surface:"#211B18", elevated:"#2E2521", text:"#F4EFE8", muted:"#9A7E76", accent:"#F5C831", accentFg:"#000000", secondary:"#C07864", border:"#3A2C29" },
  { num:"55", name:'Signal Marine', tag:'Deep Marine · Clear Telemetry', best:'Logistics, fleet, mapping tools', bg:"#111516", surface:"#1A2122", elevated:"#243031", text:"#E9F1EF", muted:"#738A87", accent:"#F5C831", accentFg:"#000000", secondary:"#36A7A0", border:"#2B3B3A" },
  { num:"56", name:'Basalt Rose', tag:'Stone Rose · Human Warmth', best:'HR systems, workplace experience apps', bg:"#151315", surface:"#211D20", elevated:"#2F272D", text:"#F4EDEE", muted:"#98838E", accent:"#F5C831", accentFg:"#000000", secondary:"#B8688C", border:"#3B3038" },
  { num:"57", name:'Blue Ash', tag:'Smoked Blue · Measured Clarity', best:'Finance dashboards, executive reporting', bg:"#121518", surface:"#1B2228", elevated:"#252F38", text:"#EEF0F2", muted:"#778490", accent:"#F5C831", accentFg:"#000000", secondary:"#6C96BC", border:"#2E3A46" },
  { num:"58", name:'Alpine Dusk', tag:'Mountain Green · Cool Evening', best:'Climate platforms, research portals', bg:"#121512", surface:"#1B211E", elevated:"#252D2A", text:"#EAF0EA", muted:"#74887C", accent:"#F5C831", accentFg:"#000000", secondary:"#7DA890", border:"#2E3A35" },
  { num:"59", name:'Ember Slate', tag:'Warm Ember · Disciplined Slate', best:'Incident tooling, security reviews', bg:"#151312", surface:"#211D1B", elevated:"#2D2825", text:"#F4EFEC", muted:"#99847B", accent:"#F5C831", accentFg:"#000000", secondary:"#D0664A", border:"#3B302C" },
  { num:"60", name:'Quantum Teal', tag:'Precise Teal · Live Systems', best:'AI operations, data observability', bg:"#101615", surface:"#182320", elevated:"#21332F", text:"#E8F1EE", muted:"#6A8E86", accent:"#F5C831", accentFg:"#000000", secondary:"#20C6B0", border:"#273F3A" },
  { num:"61", name:'Orchid Night', tag:'Soft Violet · Creative Control', best:'Creative suites, brand asset tools', bg:"#141216", surface:"#211C25", elevated:"#2E2634", text:"#F2EDF3", muted:"#97829E", accent:"#F5C831", accentFg:"#000000", secondary:"#B078D8", border:"#3A3042" },
  { num:"62", name:'Graphite Mint', tag:'Cool Mint · Clinical Graphite', best:'Medical admin, pharmacy tooling', bg:"#111514", surface:"#1A211F", elevated:"#25302C", text:"#E8F1EC", muted:"#748A80", accent:"#F5C831", accentFg:"#000000", secondary:"#66C09A", border:"#2C3B36" },
  { num:"63", name:'Dusk Clay', tag:'Muted Clay · Grounded Product', best:'Marketplaces, hospitality tools', bg:"#151312", surface:"#211E1A", elevated:"#2E2923", text:"#F3EFE8", muted:"#998A7C", accent:"#F5C831", accentFg:"#000000", secondary:"#B87858", border:"#3A312A" },
  { num:"64", name:'Polar Graphite', tag:'Cold Graphite · Scientific Calm', best:'Scientific apps, analytics workbenches', bg:"#121618", surface:"#1A2325", elevated:"#243134", text:"#EDF1F4", muted:"#74888D", accent:"#F5C831", accentFg:"#000000", secondary:"#82C8D8", border:"#2C3C40" },
  { num:"65", name:'Pine Circuit', tag:'Pine Green · Technical Focus', best:'Environmental ops, field service apps', bg:"#111512", surface:"#1A211B", elevated:"#242E25", text:"#EAF1E8", muted:"#778A76", accent:"#F5C831", accentFg:"#000000", secondary:"#58A870", border:"#2B3A2E" },
  { num:"66", name:'Noir Sand', tag:'Warm Graphite · Editorial Ease', best:'Docs, CMS, publishing dashboards', bg:"#15140F", surface:"#222015", elevated:"#312C1E", text:"#F3F0E8", muted:"#9C906E", accent:"#F5C831", accentFg:"#000000", secondary:"#D0B878", border:"#3E3623" },
  { num:"67", name:'Metro Violet', tag:'Urban Violet · Structured Motion', best:'Transport, planning, scheduling apps', bg:"#141217", surface:"#201D26", elevated:"#2C2737", text:"#F0EFF2", muted:"#8E8298", accent:"#F5C831", accentFg:"#000000", secondary:"#9868D0", border:"#383044" },
  { num:"68", name:'Copper Dusk', tag:'Copper Dusk · Premium Operations', best:'Procurement, inventory, supply chain', bg:"#15120E", surface:"#231D14", elevated:"#322719", text:"#F4EFE6", muted:"#9C8064", accent:"#F5C831", accentFg:"#000000", secondary:"#D28638", border:"#3F3020" },
  { num:"69", name:'Command Grey', tag:'Neutral Command · Enterprise Calm', best:'Admin consoles, internal platforms', bg:"#141614", surface:"#20231F", elevated:"#2B302A", text:"#F0F0EC", muted:"#879082", accent:"#F5C831", accentFg:"#000000", secondary:"#8EA084", border:"#363C34" },
  { num:"70", name:'Halo Blue', tag:'Halo Blue · Confident Signal', best:'Customer portals, SaaS control rooms', bg:"#121416", surface:"#1B2026", elevated:"#252D36", text:"#EEF0F3", muted:"#7A8492", accent:"#F5C831", accentFg:"#000000", secondary:"#5F95D6", border:"#2D3845" },
  { num:"71", name:'Hexagon', tag:'Corporate Charcoal · Brand Red Signal', best:'Banking dashboards, corporate portals', bg:"#141313", surface:"#201E1E", elevated:"#2C2A2A", text:"#F5F2F0", muted:"#9C9694", accent:"#F5C831", accentFg:"#000000", secondary:"#E8323E", border:"#383434" },
  { num:"72", name:'Lion Guard', tag:'Warm Charcoal · Heritage Red', best:'Wealth management, private banking', bg:"#151312", surface:"#221E1C", elevated:"#2F2A27", text:"#F5F0EC", muted:"#9E948C", accent:"#F5C831", accentFg:"#000000", secondary:"#D8444C", border:"#3B342F" },
  { num:"73", name:'Redline', tag:'Trading Slate · Urgent Signal', best:'Trading screens, risk monitoring', bg:"#121314", surface:"#1D1F21", elevated:"#292C2F", text:"#F0F1F2", muted:"#979CA2", accent:"#F5C831", accentFg:"#000000", secondary:"#F03A46", border:"#353A3F" },
  { num:"74", name:'Crest', tag:'Red-Warmed Charcoal · Soft Gold', best:'Brand-forward banking, flagship apps', bg:"#171214", surface:"#251B1E", elevated:"#332629", text:"#F6F0F1", muted:"#A29298", accent:"#F8D44E", accentFg:"#000000", secondary:"#E62832", border:"#412E33" },
  { num:"75", name:'Bullion', tag:'Gold-Warmed Charcoal · Vault Glow', best:'Wealth, treasury, premium financial', bg:"#161310", surface:"#242017", elevated:"#322C20", text:"#F6F1E4", muted:"#A59A7C", accent:"#F9D85C", accentFg:"#000000", secondary:"#E03038", border:"#403827" },
  { num:"76", name:'Terminal', tag:'Phosphor Black · Market Amber', best:'Trading terminals, live market data, quote screens', bg:"#0E0D0B", surface:"#181613", elevated:"#242019", text:"#F2EFE6", muted:"#9E9886", accent:"#F7A928", accentFg:"#000000", secondary:"#38C26C", border:"#342E22" },
  { num:"77", name:'Grid Slate', tag:'Navy Slate · Data Blue · Standard Yellow', best:'Data-dense grids, analytics workbenches, BI tools', bg:"#161B24", surface:"#1F2836", elevated:"#2A3547", text:"#EFF2F6", muted:"#97A4B8", accent:"#F5C831", accentFg:"#000000", secondary:"#2196F3", border:"#30405A" },
  { num:"78", name:'Dracula', tag:'Purple-Grey · Cult Classic', best:'Code editors, dev tools, terminal UIs', bg:"#21222C", surface:"#282A36", elevated:"#343746", text:"#F8F8F2", muted:"#ABADC0", accent:"#F5C831", accentFg:"#000000", secondary:"#BD93F9", border:"#3C3F52" },
  { num:"79", name:'Nord', tag:'Arctic Slate · Scandinavian Calm', best:'IDEs, documentation, focus tools', bg:"#2E3440", surface:"#3B4252", elevated:"#434C5E", text:"#ECEFF4", muted:"#B4BCCC", accent:"#F5C831", accentFg:"#000000", secondary:"#88C0D0", border:"#4C566A" },
  { num:"80", name:'Solarized', tag:'Calibrated Teal-Dark · Scientific', best:'Long-session coding, reading, terminals', bg:"#002B36", surface:"#073642", elevated:"#0E4250", text:"#E8E4D4", muted:"#9CB0B0", accent:"#F5C831", accentFg:"#000000", secondary:"#2AA198", border:"#11505F" },
  { num:"81", name:'Gruvbox', tag:'Retro Warm · Low-Fatigue', best:'Vim/terminal users, warm-tone coding', bg:"#1D2021", surface:"#282828", elevated:"#3C3836", text:"#EBDBB2", muted:"#BFB29C", accent:"#F5C831", accentFg:"#000000", secondary:"#8EC07C", border:"#423D38" },
  { num:"82", name:'Contrast Dark', tag:'Maximum Contrast · Accessibility-First', best:'Low-vision users, WCAG AAA compliance, accessibility modes', bg:"#000000", surface:"#0C0C0C", elevated:"#1A1A1A", text:"#FFFFFF", muted:"#C8C8C8", accent:"#F5C831", accentFg:"#000000", secondary:"#FFFFFF", border:"#2E2E2E" },];

const LIGHT = [
  { num:"L01", name:'Pure Canvas', tag:'Crisp · Open · Authoritative', best:'SaaS, developer tools', bg:"#F6F6F6", surface:"#FFFFFF", elevated:"#EBEBEB", text:"#080808", muted:"#606060", accent:"#B08800", accentFg:"#000000", secondary:"#101010", border:"#DCDCDC" },
  { num:"L02", name:'Warm Parchment', tag:'Inviting Warmth · Editorial Calm', best:'Documentation, blogs, reading', bg:"#FAF8F0", surface:"#FFFEF8", elevated:"#EDE8D8", text:"#16120A", muted:"#7A6C50", accent:"#A87800", accentFg:"#000000", secondary:"#786020", border:"#E2D8C0" },
  { num:"L03", name:'Frost Circuit', tag:'Cool Clarity · Tech Precision', best:'API docs, precision interfaces', bg:"#F0F4FA", surface:"#FAFCFF", elevated:"#E2ECFA", text:"#080E1C", muted:"#4E6080", accent:"#9C7800", accentFg:"#000000", secondary:"#2860C8", border:"#CCDAEC" },
  { num:"L04", name:'Sage Light', tag:'Natural Growth · Grounded Clarity', best:'Health, sustainability, agriculture', bg:"#F2F6EE", surface:"#F8FCF4", elevated:"#E4ECD8", text:"#0A0E08", muted:"#586050", accent:"#8C7200", accentFg:"#FFFFFF", secondary:"#4C7C28", border:"#CCDCC0" },
  { num:"L05", name:'Rose Forge', tag:'Bold Warmth · Human Energy', best:'Lifestyle, e-commerce, community', bg:"#FAF4F2", surface:"#FFFCFB", elevated:"#EFE0D8", text:"#160C0A", muted:"#7C5850", accent:"#986800", accentFg:"#FFFFFF", secondary:"#C83818", border:"#E4D0C8" },
  { num:"L06", name:'Pearl', tag:'Lustrous Neutral · Quiet Elegance', best:'Premium SaaS, luxury portfolio', bg:"#F8F8F5", surface:"#FFFFFF", elevated:"#EEEDE8", text:"#101010", muted:"#62645E", accent:"#A88800", accentFg:"#000000", secondary:"#505048", border:"#DCDCD4" },
  { num:"L07", name:'Arctic', tag:'Polar Blue · Crisp Intelligence', best:'Analytics, finance, research', bg:"#EEF4F8", surface:"#F8FCFF", elevated:"#DEECf8", text:"#080C14", muted:"#486080", accent:"#906800", accentFg:"#FFFFFF", secondary:"#1858A8", border:"#C8DCEC" },
  { num:"L08", name:'Linen', tag:'Fabric Warmth · Artisan Texture', best:'Publishing, craft, artisan markets', bg:"#F8F4EC", surface:"#FFFEF8", elevated:"#EEE4D0", text:"#141008", muted:"#78705A", accent:"#986800", accentFg:"#FFFFFF", secondary:"#806840", border:"#E0D4B8" },
  { num:"L09", name:'Chalk', tag:'Matte White · Soft Confidence', best:'Education, onboarding, productivity', bg:"#F4F4F2", surface:"#FAFAF8", elevated:"#E8E8E4", text:"#0C0C0A", muted:"#626260", accent:"#A08000", accentFg:"#000000", secondary:"#404038", border:"#DCDCD8" },
  { num:"L10", name:'Morning Mist', tag:'Lavender Haze · Gentle Precision', best:'Wellness, mental health, mindfulness', bg:"#F4F0F8", surface:"#FBF8FF", elevated:"#E8E0F4", text:"#0C0810", muted:"#686078", accent:"#887200", accentFg:"#FFFFFF", secondary:"#7060A0", border:"#DCD4EC" },
  { num:"L11", name:'Sand', tag:'Desert Warmth · Sun-Cured Calm', best:'Travel, hospitality, food & drink', bg:"#FAF6EC", surface:"#FFFEF4", elevated:"#EEE8D4", text:"#14100A", muted:"#786C50", accent:"#987000", accentFg:"#000000", secondary:"#907840", border:"#E4D8B8" },
  { num:"L12", name:'Seafoam', tag:'Coastal Teal · Refreshing Clarity', best:'Health tech, wellness, fintech', bg:"#EEF8F6", surface:"#F8FFFE", elevated:"#DCF0EC", text:"#081410", muted:"#447870", accent:"#7C6400", accentFg:"#FFFFFF", secondary:"#208070", border:"#C4E4DE" },
  { num:"L13", name:'Blush', tag:'Rose Quartz · Feminine Strength', best:'Beauty, fashion, lifestyle brands', bg:"#FAF0F2", surface:"#FFFBFC", elevated:"#F0E0E4", text:"#140810", muted:"#786070", accent:"#906000", accentFg:"#FFFFFF", secondary:"#C05870", border:"#E8D4D8" },
  { num:"L14", name:'Stone', tag:'Travertine · Enduring Substance', best:'Architecture, real estate, legal', bg:"#F4F2EE", surface:"#FAFAF6", elevated:"#E8E4DC", text:"#100E08", muted:"#706E64", accent:"#987200", accentFg:"#000000", secondary:"#686050", border:"#DCDAD0" },
  { num:"L15", name:'Cloud', tag:'Overcast · Soft Infrastructure', best:'SaaS dashboards, admin, B2B', bg:"#F2F4F6", surface:"#FAFCFE", elevated:"#E4E8EC", text:"#0A0C10", muted:"#606870", accent:"#907200", accentFg:"#000000", secondary:"#485870", border:"#D8DCE4" },
  { num:"L16", name:'Ecru', tag:'Raw Linen · Timeless Warmth', best:'High-end editorial, publishing', bg:"#F8F6EC", surface:"#FEFCF4", elevated:"#EEEADC", text:"#14120A", muted:"#7A7860", accent:"#A07800", accentFg:"#000000", secondary:"#807850", border:"#E4E0CC" },
  { num:"L17", name:'Mineral', tag:'Cool Slate · Crystalline Logic', best:'Engineering tools, dev portals', bg:"#EEF2F6", surface:"#F8FAFC", elevated:"#E0E8F0", text:"#0A0E14", muted:"#5A6878", accent:"#887600", accentFg:"#000000", secondary:"#406080", border:"#CCDAE8" },
  { num:"L18", name:'Ivory Slate', tag:'Warm Gray · Steady Refinement', best:'Docs, wikis, internal tools', bg:"#F6F4F0", surface:"#FEFCF8", elevated:"#ECE9E1", text:"#100E0A", muted:"#706C66", accent:"#987200", accentFg:"#000000", secondary:"#70675E", border:"#DEDAD2" },
  { num:"L19", name:'Alpine', tag:'High Altitude · Crisp & Clean', best:'Sustainability, outdoor, green tech', bg:"#EEF2F0", surface:"#F8FCF8", elevated:"#DEE8E2", text:"#08100C", muted:"#506860", accent:"#7C7000", accentFg:"#FFFFFF", secondary:"#307060", border:"#C8DCCE" },
  { num:"L20", name:'Dusk Rose', tag:'Warm Dusk · Soft Close', best:'Community, social, personal tools', bg:"#F8F0EE", surface:"#FFFBF8", elevated:"#EEDEDC", text:"#140C0A", muted:"#786860", accent:"#906800", accentFg:"#FFFFFF", secondary:"#B07068", border:"#E8D4CC" },
  { num:"L21", name:'Powder Sky', tag:'Soft Blue · Open Air', best:'Travel tech, transport, public apps', bg:"#F0F6FC", surface:"#FAFEFF", elevated:"#DEF0FA", text:"#080C14", muted:"#486080", accent:"#8C6C00", accentFg:"#FFFFFF", secondary:"#3078B8", border:"#C8E0F0" },
  { num:"L22", name:'Lavender Mist', tag:'Soft Purple · Gentle Focus', best:'Mindfulness, therapy, calm tools', bg:"#F4F0FC", surface:"#FEFBFF", elevated:"#EAE0F8", text:"#0C0814", muted:"#685888", accent:"#846400", accentFg:"#FFFFFF", secondary:"#7858B8", border:"#DDD0F0" },
  { num:"L23", name:'Mint Fresh', tag:'Crisp Mint · Clean Energy', best:'Health, fitness, nutrition apps', bg:"#EEF8F2", surface:"#F8FFF9", elevated:"#DCEFE6", text:"#081410", muted:"#4A7060", accent:"#7A6000", accentFg:"#FFFFFF", secondary:"#2A9060", border:"#C0E4CC" },
  { num:"L24", name:'Honey', tag:'Warm Gold · Nourishing Clarity', best:'Food tech, recipe, wellness commerce', bg:"#FAF4E4", surface:"#FFFCF0", elevated:"#F0E8CC", text:"#14100A", muted:"#786848", accent:"#A07400", accentFg:"#000000", secondary:"#C09028", border:"#E8D8A8" },
  { num:"L25", name:'Coral Reef', tag:'Warm Coral · Living Energy', best:'Social, creator, community apps', bg:"#FEF4F0", surface:"#FFFBF8", elevated:"#FAE4DC", text:"#180C08", muted:"#806058", accent:"#9C5C00", accentFg:"#FFFFFF", secondary:"#D05030", border:"#EED4C8" },
  { num:"L26", name:'Sage Mist', tag:'Pale Sage · Quiet Growth', best:'Organic brands, plant-based commerce', bg:"#F0F6F2", surface:"#F8FEF8", elevated:"#E0EEE4", text:"#0A100C", muted:"#527060", accent:"#7C6800", accentFg:"#FFFFFF", secondary:"#4A8A54", border:"#C8DECC" },
  { num:"L27", name:'Ice Sheet', tag:'Polar White · Arctic Clarity', best:'Scientific tools, pharma, cold-chain', bg:"#EDF5F8", surface:"#F8FDFF", elevated:"#D9EEF6", text:"#080E14", muted:"#4C6A78", accent:"#8A7000", accentFg:"#FFFFFF", secondary:"#48A8BC", border:"#C8E0E8" },
  { num:"L28", name:'Amber Glow', tag:'Warm Amber · Sunset Logic', best:'Energy platforms, solar tech', bg:"#FDF6E8", surface:"#FFFEF2", elevated:"#F8ECC8", text:"#140E04", muted:"#7A6840", accent:"#A07200", accentFg:"#000000", secondary:"#C89020", border:"#EEE0A8" },
  { num:"L29", name:'Lilac', tag:'Soft Lilac · Gentle Precision', best:'Education, student, learning tools', bg:"#F6F0FC", surface:"#FDF8FF", elevated:"#ECE0F8", text:"#100814", muted:"#706880", accent:"#806000", accentFg:"#FFFFFF", secondary:"#9068C0", border:"#E0CCED" },
  { num:"L30", name:'Warm Dove', tag:'Warm Grey · Considered Calm', best:'Consulting, professional services', bg:"#F4F2EF", surface:"#FEFCFA", elevated:"#EAE5DE", text:"#100E0C", muted:"#6C665E", accent:"#987000", accentFg:"#000000", secondary:"#82786A", border:"#DED8CE" },
  { num:"L31", name:'Sage White', tag:'Near White · Botanical Quiet', best:'Spa, clean beauty, holistic health', bg:"#F2F6F0", surface:"#FAFEF8", elevated:"#E2EEE0", text:"#0A0E08", muted:"#587858", accent:"#7A6400", accentFg:"#FFFFFF", secondary:"#508048", border:"#CCDEC8" },
  { num:"L32", name:'Steel Blue Light', tag:'Pale Steel · Confident Clarity', best:'Procurement, supply chain, corporate', bg:"#EEF2F8", surface:"#F8FAFF", elevated:"#DFEAF8", text:"#080C14", muted:"#506080", accent:"#8C7800", accentFg:"#000000", secondary:"#4070A8", border:"#CAD8EA" },
  { num:"L33", name:'Terracotta Light', tag:'Warm Clay · Mediterranean Sun', best:'Hospitality, food & beverage, culture', bg:"#F8EFE8", surface:"#FFFAF4", elevated:"#F0DECC", text:"#140C08", muted:"#80604E", accent:"#9A6000", accentFg:"#FFFFFF", secondary:"#BC603C", border:"#E8CEB6" },
  { num:"L34", name:'Pearl Blue', tag:'Iridescent Pearl · Premium Tech', best:'Luxury SaaS, high-end e-commerce', bg:"#F1F5FA", surface:"#FAFDFF", elevated:"#E3ECF6", text:"#080C14", muted:"#526272", accent:"#906800", accentFg:"#FFFFFF", secondary:"#4A7FB2", border:"#CDDCEB" },
  { num:"L35", name:'Wheat', tag:'Sun-Dried Grain · Rural Clarity', best:'AgriTech, rural services, commodities', bg:"#F8F4E8", surface:"#FFFEF0", elevated:"#EEE8D0", text:"#14100A", muted:"#7A7050", accent:"#A07800", accentFg:"#000000", secondary:"#9C8840", border:"#E4DAB8" },
  { num:"L36", name:'Aqua Fresh', tag:'Pale Aqua · Refreshed Mind', best:'Hydration apps, swimming tech', bg:"#EEF8F8", surface:"#F8FFFF", elevated:"#DCEEEE", text:"#081212", muted:"#4A6868", accent:"#7C6800", accentFg:"#FFFFFF", secondary:"#2898A0", border:"#C4E0E0" },
  { num:"L37", name:'Rose Gold', tag:'Blush Metal · Warm Aspiration', best:'Jewellery, accessories, lifestyle', bg:"#FBF1F3", surface:"#FFFBFC", elevated:"#F3E0E5", text:"#140A0C", muted:"#80636A", accent:"#9A6200", accentFg:"#FFFFFF", secondary:"#C87888", border:"#EDD6DC" },
  { num:"L38", name:'Dove Grey', tag:'Soft Dove · Understated Elegance', best:'Law firms, accounting, professional', bg:"#F4F4F2", surface:"#FEFEFE", elevated:"#EAEAE6", text:"#0E0E0C", muted:"#6A6A68", accent:"#9A7600", accentFg:"#000000", secondary:"#606058", border:"#DCDCD8" },
  { num:"L39", name:'Peach', tag:'Warm Peach · Approachable Energy', best:'Consumer apps, family tech, parenting', bg:"#FDF4EE", surface:"#FFFAF6", elevated:"#F8E4D4", text:"#180C08", muted:"#806458", accent:"#9A6000", accentFg:"#FFFFFF", secondary:"#D07850", border:"#EED8C8" },
  { num:"L40", name:'Graphite White', tag:'Near White · Silent Authority', best:'White-label tools, neutral admin', bg:"#F2F2EF", surface:"#FCFCFA", elevated:"#E7E7E1", text:"#0E0E0C", muted:"#686864", accent:"#9A7800", accentFg:"#000000", secondary:"#5C625C", border:"#DADAD2" },
  { num:"L41", name:'Spring Blossom', tag:'Fresh Green · New Growth', best:'Garden tech, horticulture, farming', bg:"#EEF8EE", surface:"#F8FFF8", elevated:"#DEF0DE", text:"#081408", muted:"#507A50", accent:"#786C00", accentFg:"#FFFFFF", secondary:"#489848", border:"#C8E8C8" },
  { num:"L42", name:'Horizon', tag:'Pale Horizon · Open Distance', best:'Aviation, navigation, planning tools', bg:"#EEF6FC", surface:"#F8FCFF", elevated:"#DCEBFA", text:"#080C12", muted:"#486078", accent:"#8A7200", accentFg:"#FFFFFF", secondary:"#5AA0D8", border:"#C8DCF0" },
  { num:"L43", name:'Buttermilk', tag:'Warm Cream · Nourishing Light', best:'Recipe platforms, home cooking', bg:"#FAF6E8", surface:"#FFFEF2", elevated:"#F4ECCC", text:"#141008", muted:"#7A7848", accent:"#A07600", accentFg:"#000000", secondary:"#B09430", border:"#ECE4B0" },
  { num:"L44", name:'Clay', tag:'Hand-Thrown Clay · Craft Authority', best:'Pottery, craft, artisan design tools', bg:"#F6F0EB", surface:"#FEFBF6", elevated:"#EDE3D8", text:"#12100C", muted:"#726058", accent:"#986400", accentFg:"#FFFFFF", secondary:"#9C7060", border:"#E0D0C0" },
  { num:"L45", name:'Powder Pink', tag:'Delicate Pink · Soft Strength', best:'Personal care, maternity, wellness', bg:"#FDF3F7", surface:"#FFFAFD", elevated:"#F8E3EC", text:"#140A0E", muted:"#806072", accent:"#906000", accentFg:"#FFFFFF", secondary:"#D08098", border:"#F0D8E2" },
  { num:"L46", name:'Cool Mint', tag:'Crisp Mint · Alert Calm', best:'Pharmacy, clinical, medical SaaS', bg:"#ECF8F4", surface:"#F6FFFC", elevated:"#D8F0E8", text:"#081410", muted:"#488068", accent:"#726800", accentFg:"#FFFFFF", secondary:"#30A880", border:"#C0E8DC" },
  { num:"L47", name:'Warm Cream', tag:'Pure Cream · Timeless Ground', best:'Luxury retail, premium memberships', bg:"#FBF8F0", surface:"#FFFEFB", elevated:"#F4EDD8", text:"#14100A", muted:"#7A7458", accent:"#9C7A00", accentFg:"#000000", secondary:"#8C7C48", border:"#EAE0C0" },
  { num:"L48", name:'Slate Violet', tag:'Pale Violet · Structured Dream', best:'Creative agencies, branding tools', bg:"#F0F0F8", surface:"#FAFAFF", elevated:"#E4E4F4", text:"#0C0C14", muted:"#686878", accent:"#7E6C00", accentFg:"#FFFFFF", secondary:"#7870C0", border:"#D8D8F0" },
  { num:"L49", name:'Cotton', tag:'Pure Cotton · Absolute Softness', best:'Baby & child platforms, family apps', bg:"#FAFAF7", surface:"#FFFFFF", elevated:"#F0F0EA", text:"#0C0C0A", muted:"#6C6C66", accent:"#9A7800", accentFg:"#000000", secondary:"#7A7566", border:"#E2E2D8" },
  { num:"L50", name:'Champagne', tag:'Effervescent Gold · Celebratory Clarity', best:'Events tech, ticketing, experiences', bg:"#FAF6EC", surface:"#FFFEF4", elevated:"#F4ECD4", text:"#14100A", muted:"#807660", accent:"#A07A00", accentFg:"#000000", secondary:"#C0A840", border:"#EAE0BC" },
  { num:"L51", name:'Porcelain Circuit', tag:'Porcelain White · Fine Signal', best:'Premium SaaS, executive portals', bg:"#F7F7F4", surface:"#FFFFFF", elevated:"#ECECE6", text:"#10100C", muted:"#686860", accent:"#987200", accentFg:"#000000", secondary:"#3C4A58", border:"#DDDCD4" },
  { num:"L52", name:'Mist Graphite', tag:'Soft Grey · Quiet Authority', best:'Admin tools, enterprise back-office', bg:"#F0F3F0", surface:"#FBFCFA", elevated:"#E2E8E2", text:"#0E100E", muted:"#5E6C60", accent:"#947400", accentFg:"#000000", secondary:"#72806C", border:"#D0D8D0" },
  { num:"L53", name:'Glacier Mint', tag:'Icy Mint · Clinical Clarity', best:'Healthcare, lab operations, pharma', bg:"#EDF8F6", surface:"#F8FFFD", elevated:"#DAF0EC", text:"#081410", muted:"#4A746A", accent:"#786800", accentFg:"#FFFFFF", secondary:"#28B092", border:"#C2E6DC" },
  { num:"L54", name:'Soft Copper', tag:'Pale Copper · Crafted Systems', best:'Hardware, manufacturing, maker tools', bg:"#F8F0E7", surface:"#FFFBF4", elevated:"#EEDBCC", text:"#140D08", muted:"#7C6252", accent:"#986400", accentFg:"#FFFFFF", secondary:"#C4865A", border:"#E2CCB8" },
  { num:"L55", name:'Cloud Violet', tag:'Cloud Violet · Creative Calm', best:'Design systems, creative workflow apps', bg:"#F3F1FA", surface:"#FCFAFF", elevated:"#E5E0F4", text:"#0E0B14", muted:"#686078", accent:"#806800", accentFg:"#FFFFFF", secondary:"#8A70C8", border:"#D8D0EC" },
  { num:"L56", name:'Paper Sage', tag:'Paper Sage · Natural Order', best:'Sustainability, HR, planning tools', bg:"#F0F5EE", surface:"#FAFEF7", elevated:"#E0EADC", text:"#0C100A", muted:"#5B7154", accent:"#827000", accentFg:"#FFFFFF", secondary:"#638A52", border:"#CBDCC4" },
  { num:"L57", name:'Pearl Coral', tag:'Pearl Coral · Human Product', best:'Consumer apps, onboarding, community', bg:"#FCF2EF", surface:"#FFFBF9", elevated:"#F3E1DC", text:"#160B08", muted:"#7A5F58", accent:"#965E00", accentFg:"#FFFFFF", secondary:"#C86654", border:"#EAD3CC" },
  { num:"L58", name:'Blue Linen', tag:'Blue Linen · Business Clarity', best:'Finance, analytics, CRM tools', bg:"#F0F3F7", surface:"#FBFCFF", elevated:"#E2E8F0", text:"#080D14", muted:"#586678", accent:"#8C7000", accentFg:"#FFFFFF", secondary:"#587FA8", border:"#D0D9E4" },
  { num:"L59", name:'Cream Graphite', tag:'Cream Ground · Sharp Structure', best:'Documentation, consulting, legal tech', bg:"#F8F5ED", surface:"#FFFDF6", elevated:"#EEE6D6", text:"#12100A", muted:"#746A58", accent:"#987200", accentFg:"#000000", secondary:"#5A5A52", border:"#E0D7C4" },
  { num:"L60", name:'Solar Mist', tag:'Solar Mist · Warm Intelligence', best:'Energy, climate, operations dashboards', bg:"#FAF4E2", surface:"#FFFDF0", elevated:"#F3E5BC", text:"#140E04", muted:"#7A6638", accent:"#A07400", accentFg:"#000000", secondary:"#D0A018", border:"#E8D69C" },
  { num:"L61", name:'Pale Lagoon', tag:'Pale Lagoon · Clean Depth', best:'Marine, logistics, water platforms', bg:"#ECF8F6", surface:"#F8FFFE", elevated:"#D8EEE8", text:"#081412", muted:"#48746E", accent:"#7C6600", accentFg:"#FFFFFF", secondary:"#24948A", border:"#C2E2DC" },
  { num:"L62", name:'Chalk Plum', tag:'Chalk Plum · Soft Focus', best:'Education, research, writing tools', bg:"#F6F2F8", surface:"#FFFBFF", elevated:"#ECE2F0", text:"#120C14", muted:"#74647C", accent:"#846400", accentFg:"#FFFFFF", secondary:"#9868A8", border:"#E0D0E8" },
  { num:"L63", name:'Studio White', tag:'Studio White · Gallery Precision', best:'Creative portfolios, asset libraries', bg:"#F6F6F4", surface:"#FFFFFF", elevated:"#ECECEA", text:"#0E0E0C", muted:"#666664", accent:"#967600", accentFg:"#000000", secondary:"#303030", border:"#DEDEDA" },
  { num:"L64", name:'Quiet Rose', tag:'Quiet Rose · Considered Warmth', best:'Wellness, workplace, personal tools', bg:"#FAF0F2", surface:"#FFFBFC", elevated:"#F0DEE3", text:"#140A0C", muted:"#7A5E68", accent:"#906200", accentFg:"#FFFFFF", secondary:"#B06A7C", border:"#E8D0D8" },
  { num:"L65", name:'Mineral Green', tag:'Mineral Green · Practical Freshness', best:'Field ops, agriculture, inventory', bg:"#EDF5F1", surface:"#F8FEFA", elevated:"#D9E9E1", text:"#09120C", muted:"#4C705C", accent:"#7E6C00", accentFg:"#FFFFFF", secondary:"#2F9270", border:"#C4DACB" },
  { num:"L66", name:'Frost Amber', tag:'Frost Amber · Clean Energy', best:'Energy SaaS, billing, reporting', bg:"#F6F4EC", surface:"#FFFEF8", elevated:"#ECE8D8", text:"#12100A", muted:"#706850", accent:"#9A7600", accentFg:"#000000", secondary:"#B89030", border:"#E0DAC4" },
  { num:"L67", name:'Porcelain Navy', tag:'Porcelain Navy · Executive Calm', best:'Fintech, governance, board portals', bg:"#F0F4F8", surface:"#FAFDFF", elevated:"#E0E8F0", text:"#080C14", muted:"#4E5E70", accent:"#907000", accentFg:"#FFFFFF", secondary:"#304C78", border:"#CCD8E6" },
  { num:"L68", name:'Warm Steel', tag:'Warm Steel · Industrial Light', best:'Procurement, supply chain, logistics', bg:"#F3F2EF", surface:"#FEFCF8", elevated:"#E6E4DE", text:"#100E0A", muted:"#6A6860", accent:"#967200", accentFg:"#000000", secondary:"#687078", border:"#DAD6CE" },
  { num:"L69", name:'Clinical Pearl', tag:'Clinical Pearl · Exact Care', best:'Clinical dashboards, patient ops', bg:"#F1F7F6", surface:"#FAFFFE", elevated:"#E0EEEC", text:"#081210", muted:"#4E6F6A", accent:"#766A00", accentFg:"#FFFFFF", secondary:"#4A98A0", border:"#C8E0DC" },
  { num:"L70", name:'Horizon Gold', tag:'Horizon Gold · Optimistic Precision', best:'Travel, planning, customer portals', bg:"#F6F4EA", surface:"#FFFDF4", elevated:"#ECE7D2", text:"#121008", muted:"#746C50", accent:"#9A7800", accentFg:"#000000", secondary:"#B89A3C", border:"#E2D8B8" },
  { num:"L71", name:'Hexagon White', tag:'Brand White · True Red', best:'Customer-facing banking, public sites', bg:"#F6F5F4", surface:"#FFFFFF", elevated:"#ECEAE8", text:"#0E0C0C", muted:"#5E5A58", accent:"#967000", accentFg:"#FFFFFF", secondary:"#DB0011", border:"#DEDAD8" },
  { num:"L72", name:'Silk Red', tag:'Warm Porcelain · Deep Red', best:'Premier banking, relationship tools', bg:"#FAF5F4", surface:"#FFFBFA", elevated:"#F2E6E4", text:"#140C0C", muted:"#6E5C5A", accent:"#9A6800", accentFg:"#FFFFFF", secondary:"#C00010", border:"#E8D6D4" },
  { num:"L73", name:'City Grey', tag:'Corporate Grey · Precise Red', best:'Internal corporate tools, compliance', bg:"#F2F3F4", surface:"#FAFBFC", elevated:"#E4E6E8", text:"#0C0D0E", muted:"#5A6066", accent:"#8C7000", accentFg:"#FFFFFF", secondary:"#DB0011", border:"#D6DADE" },
  { num:"L74", name:'Ivory Crest', tag:'Warm Ivory · Bright Gold · True Red', best:'Flagship customer apps, brand sites', bg:"#F8F4F2", surface:"#FFFDFC", elevated:"#F0E4E0", text:"#140D0C", muted:"#6E5A56", accent:"#A88408", accentFg:"#000000", secondary:"#DB0011", border:"#E6D4CE" },
  { num:"L75", name:'Gilt', tag:'Gold-Washed White · Gilt-Edged', best:'Premium wealth, investor portals', bg:"#F9F5E9", surface:"#FFFDF3", elevated:"#F2EAD0", text:"#141106", muted:"#6F684C", accent:"#A8860C", accentFg:"#000000", secondary:"#C00010", border:"#E6DDBD" },
  { num:"L76", name:'Quartz Light', tag:'Cold White · Grid Blue', best:'Spreadsheet apps, reporting tools, admin grids', bg:"#F4F6F8", surface:"#FCFDFE", elevated:"#E6EAEE", text:"#0C0E10", muted:"#5A626C", accent:"#8C7000", accentFg:"#FFFFFF", secondary:"#1E88E5", border:"#D8DDE3" },
  { num:"L77", name:'Ticker', tag:'Ticker-Tape Paper · Market Green', best:'Market summaries, investor reports, daylight trading', bg:"#FAF7EE", surface:"#FFFDF6", elevated:"#F0EBD8", text:"#13110A", muted:"#6E6850", accent:"#A88000", accentFg:"#000000", secondary:"#1E8A4C", border:"#E2DCC4" },
  { num:"L78", name:'Contrast Light', tag:'Maximum Contrast · Accessibility-First', best:'Low-vision users, WCAG AAA compliance, accessibility modes', bg:"#FFFFFF", surface:"#FFFFFF", elevated:"#F0F0F0", text:"#000000", muted:"#3A3A3A", accent:"#7A5C00", accentFg:"#FFFFFF", secondary:"#000000", border:"#B0B0B0" },];

function CircuitMark({ size = 32, color = "#F5C831" }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none">
      <rect x="1.5" y="1.5" width="29" height="29" rx="3.5" stroke={color} strokeWidth="1.6"/>
      <circle cx="1.5" cy="1.5" r="2.5" fill={color}/><circle cx="30.5" cy="1.5" r="2.5" fill={color}/>
      <circle cx="1.5" cy="30.5" r="2.5" fill={color}/><circle cx="30.5" cy="30.5" r="2.5" fill={color}/>
      <circle cx="16" cy="16" r="2" fill={color}/>
      <line x1="16" y1="16" x2="22" y2="16" stroke={color} strokeWidth="1.4"/>
      <circle cx="22" cy="16" r="1.2" fill={color} fillOpacity="0.5"/>
      <line x1="16" y1="16" x2="16" y2="22" stroke={color} strokeWidth="1.4"/>
      <circle cx="16" cy="22" r="1.2" fill={color} fillOpacity="0.5"/>
      <line x1="16" y1="16" x2="10" y2="16" stroke={color} strokeWidth="1.4"/>
      <circle cx="10" cy="16" r="1.2" fill={color} fillOpacity="0.3"/>
      <line x1="16" y1="16" x2="16" y2="10" stroke={color} strokeWidth="1.4"/>
      <circle cx="16" cy="10" r="1.2" fill={color} fillOpacity="0.3"/>
    </svg>
  );
}

function PaletteCard({ p, onClick }) {
  const [h, setH] = useState(false);
  const tk = ["bg","surface","elevated","text","muted","accent","secondary","border"];
  return (
    <div onClick={() => onClick(p)} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{ background:p.surface, border:`1px solid ${h?p.accent:p.border}`, borderRadius:9, overflow:"hidden", cursor:"pointer", transition:"all 0.16s ease", transform:h?"translateY(-3px)":"none", boxShadow:h?`0 10px 32px ${p.accent}26`:"none" }}>
      <div style={{ display:"flex", height:5 }}>{tk.map(k => <div key={k} style={{ flex:1, background:p[k] }}/>)}</div>
      <div style={{ background:p.bg, padding:"11px 11px 9px" }}>
        <div style={{ background:p.surface, borderRadius:5, border:`1px solid ${p.border}`, overflow:"hidden", marginBottom:8 }}>
          <div style={{ background:p.elevated, padding:"4px 7px", display:"flex", alignItems:"center", justifyContent:"space-between", borderBottom:`1px solid ${p.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:4 }}>
              <div style={{ width:7, height:7, borderRadius:1, background:p.accent }}/>
              <span style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:8, color:p.text, letterSpacing:"0.12em" }}>FRAME</span>
            </div>
            <div style={{ display:"flex", gap:2 }}>{[p.accent,p.secondary,p.muted].map((c,i)=><div key={i} style={{ width:4, height:4, borderRadius:"50%", background:c, opacity:i===0?1:0.5 }}/>)}</div>
          </div>
          <div style={{ padding:"6px" }}>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:11, color:p.accent, letterSpacing:"0.05em", marginBottom:1 }}>24</div>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:300, fontSize:7.5, color:p.muted, marginBottom:5 }}>Active apps</div>
            <div style={{ display:"flex", gap:2 }}>
              <div style={{ flex:1, height:2.5, background:p.accent, borderRadius:2 }}/>
              <div style={{ flex:1, height:2.5, background:`${p.secondary}66`, borderRadius:2 }}/>
              <div style={{ flex:1, height:2.5, background:`${p.muted}44`, borderRadius:2 }}/>
            </div>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:4 }}>
          <div style={{ flex:1, minWidth:0 }}>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:10.5, color:p.text, letterSpacing:"0.04em", marginBottom:1, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{p.name}</div>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:300, fontSize:8, color:p.muted, lineHeight:1.3, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{p.tag}</div>
          </div>
          <span style={{ fontFamily:"'Space Mono',monospace", fontWeight:700, fontSize:7.5, padding:"1px 5px", borderRadius:3, letterSpacing:"0.05em", background:`${p.accent}20`, color:p.accent, flexShrink:0 }}>{p.num}</span>
        </div>
      </div>
    </div>
  );
}

function Modal({ p, onClose }) {
  if (!p) return null;
  const tk=[["bg","Background"],["surface","Surface"],["elevated","Elevated"],["text","Text"],["muted","Muted"],["accent","Accent"],["accentFg","Accent FG"],["secondary","Secondary"],["border","Border"]];
  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.85)", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div onClick={e=>e.stopPropagation()} style={{ background:p.surface, border:`1px solid ${p.border}`, borderRadius:14, width:"100%", maxWidth:500, overflow:"hidden", maxHeight:"90vh", overflowY:"auto" }}>
        <div style={{ background:p.bg, padding:"20px 24px 16px", borderBottom:`1px solid ${p.border}` }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:9 }}>
            <div style={{ width:9, height:9, borderRadius:2, background:p.accent }}/>
            <span style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:10, color:p.accent, letterSpacing:"0.16em" }}>FRAME PALETTE · {p.num}</span>
          </div>
          <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:26, color:p.text, letterSpacing:"0.03em", marginBottom:3 }}>{p.name}</div>
          <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:300, fontSize:12, color:p.muted }}>{p.tag}</div>
        </div>
        <div style={{ height:8, display:"flex" }}>{tk.map(([k]) => <div key={k} style={{ flex:1, background:p[k] }}/>)}</div>
        <div style={{ padding:"16px 24px" }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:12 }}>
            {tk.map(([k,label]) => (
              <div key={k} style={{ display:"flex", alignItems:"center", gap:7, padding:"5px 8px", background:p.bg, borderRadius:5, border:`1px solid ${p.border}` }}>
                <div style={{ width:17, height:17, borderRadius:2, background:p[k], border:`1px solid ${p.border}`, flexShrink:0 }}/>
                <div><div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:600, fontSize:7.5, color:p.muted, letterSpacing:"0.1em" }}>{label.toUpperCase()}</div>
                <div style={{ fontFamily:"'Space Mono',monospace", fontSize:9.5, color:p.text }}>{p[k]}</div></div>
              </div>
            ))}
          </div>
          <div style={{ padding:"10px 12px", background:p.bg, borderRadius:7, border:`1px solid ${p.border}`, marginBottom:12 }}>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:600, fontSize:8.5, color:p.accent, letterSpacing:"0.14em", marginBottom:3 }}>BEST FOR</div>
            <div style={{ fontFamily:"'Chakra Petch',sans-serif", fontWeight:400, fontSize:12, color:p.text }}>{p.best}</div>
          </div>
          <button onClick={onClose} style={{ width:"100%", background:p.accent, color:p.accentFg, border:"none", borderRadius:7, padding:"9px", fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:11, letterSpacing:"0.1em", cursor:"pointer" }}>CLOSE</button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const [sel, setSel] = useState(null);

  const fd = useMemo(() => DARK.filter(p => !q || p.name.toLowerCase().includes(q.toLowerCase()) || p.tag.toLowerCase().includes(q.toLowerCase())), [q]);
  const fl = useMemo(() => LIGHT.filter(p => !q || p.name.toLowerCase().includes(q.toLowerCase()) || p.tag.toLowerCase().includes(q.toLowerCase())), [q]);

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Chakra+Petch:wght@300;400;600;700&family=Space+Mono&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        ::-webkit-scrollbar { width:4px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:#222; border-radius:2px; }
      `}</style>
      <div style={{ background:"#0C0C0B", minHeight:"100vh", color:"#fff", fontFamily:"'Chakra Petch',sans-serif" }}>
        <div style={{ textAlign:"center", padding:"50px 24px 36px", background:"radial-gradient(ellipse 80% 60% at 50% 0%, #1A1500 0%, #0C0C0B 70%)", borderBottom:"1px solid #1A1A18" }}>
          <div style={{ display:"flex", justifyContent:"center", marginBottom:16 }}><CircuitMark size={44} color="#F5C831"/></div>
          <div style={{ fontWeight:300, fontSize:10, letterSpacing:"0.4em", color:"#F5C831", marginBottom:13, textTransform:"uppercase" }}>The FRAME Design System</div>
          <h1 style={{ fontWeight:700, fontSize:46, letterSpacing:"0.03em", color:"#FFF", lineHeight:1.05, marginBottom:13 }}>160 Palettes.<br/><span style={{ color:"#F5C831" }}>One ecosystem.</span></h1>
          <p style={{ fontWeight:300, fontSize:15, color:"#666", maxWidth:500, margin:"0 auto 28px", lineHeight:1.7 }}>Every FRAME app speaks the same visual language — anchored by one signature yellow across 160 precision-crafted, WCAG-verified palettes.</p>
          <div style={{ display:"inline-flex", gap:1, background:"#141414", border:"1px solid #1E1E1E", borderRadius:10, overflow:"hidden", marginBottom:26 }}>
            {[["160","Palettes"],["82","Dark"],["78","Light"],["1","Accent"]].map(([n,l])=>(
              <div key={l} style={{ padding:"10px 22px", borderRight:"1px solid #1E1E1E" }}>
                <div style={{ fontWeight:700, fontSize:20, color:"#F5C831" }}>{n}</div>
                <div style={{ fontWeight:400, fontSize:9, color:"#555", letterSpacing:"0.1em", textTransform:"uppercase" }}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{ display:"flex", justifyContent:"center", gap:10, alignItems:"center", flexWrap:"wrap" }}>
            <div style={{ display:"flex", gap:2, background:"#121212", border:"1px solid #1E1E1E", borderRadius:8, padding:3 }}>
              {[["all","All 160"],["dark","82 Dark"],["light","78 Light"]].map(([v,l])=>(
                <button key={v} onClick={()=>setFilter(v)} style={{ padding:"7px 16px", borderRadius:6, border:"none", cursor:"pointer", background:filter===v?"#F5C831":"transparent", color:filter===v?"#000":"#555", fontFamily:"'Chakra Petch',sans-serif", fontWeight:700, fontSize:11, letterSpacing:"0.06em", transition:"all 0.15s" }}>{l}</button>
              ))}
            </div>
            <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search palettes…" style={{ background:"#121212", border:"1px solid #1E1E1E", borderRadius:8, padding:"8px 14px", color:"#FFF", fontFamily:"'Chakra Petch',sans-serif", fontSize:11, letterSpacing:"0.04em", outline:"none", width:180 }}/>
          </div>
        </div>

        {filter!=="light" && (<>
          <div style={{ padding:"26px 26px 12px" }}><div style={{ display:"flex", alignItems:"center", gap:11 }}><div style={{ width:3, height:18, background:"#F5C831", borderRadius:2 }}/><span style={{ fontWeight:700, fontSize:10, color:"#F5C831", letterSpacing:"0.2em" }}>DARK · 01–82 {q&&`· ${fd.length} match`}</span></div></div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(178px,1fr))", gap:10, padding:"0 26px 26px" }}>{fd.map(p => <PaletteCard key={p.num} p={p} onClick={setSel}/>)}</div>
        </>)}
        {filter!=="dark" && (<>
          <div style={{ padding:"26px 26px 12px", borderTop:filter==="all"?"1px solid #141414":"none" }}><div style={{ display:"flex", alignItems:"center", gap:11 }}><div style={{ width:3, height:18, background:"#A87800", borderRadius:2 }}/><span style={{ fontWeight:700, fontSize:10, color:"#A87800", letterSpacing:"0.2em" }}>LIGHT · L01–L78 {q&&`· ${fl.length} match`}</span></div></div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(178px,1fr))", gap:10, padding:"0 26px 40px" }}>{fl.map(p => <PaletteCard key={p.num} p={p} onClick={setSel}/>)}</div>
        </>)}

        <div style={{ borderTop:"1px solid #141414", padding:"24px 26px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}><CircuitMark size={18} color="#F5C831"/><span style={{ fontWeight:700, fontSize:12, letterSpacing:"0.14em" }}>FRAME</span></div>
          <span style={{ fontWeight:300, fontSize:10, color:"#333", letterSpacing:"0.1em" }}>160 PALETTES · ONE ACCENT · ALL WCAG VERIFIED</span>
          <div style={{ fontFamily:"'Space Mono',monospace", fontSize:10, color:"#F5C831" }}>#F5C831</div>
        </div>
      </div>
      <Modal p={sel} onClose={()=>setSel(null)}/>
    </>
  );
}
