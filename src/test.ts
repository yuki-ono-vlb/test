function test_seve(){
	let name = "test";
	let date = setDayjs().format(DATE_FORMAT);
	let comment = "本日は晴天なり";
	seve(name,date,false,comment);
}

function test_getMember(){
	let filters = [
		"", // 名前
		"新井", // 担当
		"", // 会社
		"", // 所属課
	]
	const result = getMember(filters);
	Logger.log(result)
}