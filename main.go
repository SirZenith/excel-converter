package main

import (
	"context"
	"os"
	"time"

	"github.com/charmbracelet/log"
	"github.com/urfave/cli/v3"
)

const keyEnvLogLevel = "EXCEL_CONVERTER_LOG_LEVEL"

func main() {
	logger := log.NewWithOptions(os.Stderr, log.Options{
		ReportTimestamp: true,
		TimeFormat:      time.TimeOnly,
	})
	log.SetDefault(logger)
	log.SetLevel(getLogLevel())

	cmd := &cli.Command{
		Name:                  "excel-converter",
		Usage:                 "a tool that converts Excel file to JSON",
		Version:               "0.1.0",
		EnableShellCompletion: true,
		Commands: []*cli.Command{
			cmdToJSON(),
			cmdToScript(),
		},
	}

	err := cmd.Run(context.Background(), os.Args)
	if err != nil {
		log.Error(err)
		os.Exit(1)
	}
}
