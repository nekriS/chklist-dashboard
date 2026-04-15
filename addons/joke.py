import requests

payload = {
	"model": "qwen3:14b",
	"prompt": "Напиши смешной короткий анекдот на 2 предложения",
	"stream": False,
	"think": False
}

def getJoke(config, promt):

	url = config["JOKE"]["URL"]
	payload = {
		"model": "qwen3:14b",
		"prompt": f"{promt}",
		"stream": False,
		"think": False
	}
	try:
		response = requests.post(url, json=payload)
		return(response.json()['response'])
	except:
		return "Генератор шуток сломался, пинайте 18-ого..."