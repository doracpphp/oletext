// Command oletext extracts text from legacy Office binary files
// (.doc, .xls, .ppt) and writes it to standard output as UTF-8.
package main

import (
	"fmt"
	"os"

	"oletext"
)

func main() {
	args := os.Args[1:]
	if len(args) == 0 {
		fmt.Fprintln(os.Stderr, "usage: oletext <file.doc|file.xls|file.ppt> ...")
		os.Exit(2)
	}
	exitCode := 0
	for _, path := range args {
		text, err := oletext.ExtractFile(path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "oletext: %s: %v\n", path, err)
			exitCode = 1
			continue
		}
		if len(args) > 1 {
			fmt.Printf("==> %s <==\n", path)
		}
		fmt.Print(text)
		if len(args) > 1 {
			fmt.Println()
		}
	}
	os.Exit(exitCode)
}
