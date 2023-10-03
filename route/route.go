package route

import (
	"pivot/controller"

	"github.com/gin-gonic/gin"
)

func SetupRouter() *gin.Engine {
	r := gin.Default()
	r.GET("/pivoteTable/:fn", controller.GetPivotTable)
	r.POST("/upload", controller.UploadFile)

	r.Run(":8080")
	return r

}
