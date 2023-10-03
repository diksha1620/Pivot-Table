package main

import (
	"os"
	"pivot/route"

	"github.com/joho/godotenv"
)

func main() {

	godotenv.Load() // Load env variables
	//models.ConnectToDatabase() // load db
	router := route.SetupRouter()

	port := os.Getenv("port")

	if port == "" {
		port = "8080"
	}

	router.Run(":" + port)

}
