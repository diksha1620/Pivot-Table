package controller

import (
	"fmt"
	"net/http"
	"os"
	"pivot/services"

	"github.com/gin-gonic/gin"
)

func GetPivotTable(c *gin.Context) {

	var filename string = "uploads/" + c.Param("fn") + ".xlsx"
	services.CreatePivoteTable(filename, c)

}

// UploadFile handles the file upload via a POST request
func UploadFile(c *gin.Context) {
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	// Create upload directory if it doesn't exist
	uploadDir := "uploads"
	if _, err := os.Stat(uploadDir); os.IsNotExist(err) {
		err := os.Mkdir(uploadDir, os.ModePerm)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create upload directory"})
			return
		}
	}

	// Generate a unique file name for the uploaded file
	fileName := fmt.Sprintf("%s/%s%s", uploadDir, "uploaded_", file.Filename)

	// Save the uploaded file to the specified path
	if err := c.SaveUploadedFile(file, fileName); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to save file"})
		return
	}

	c.JSON(http.StatusOK, gin.H{"message": "File uploaded successfully", "file_path": fileName})
}
