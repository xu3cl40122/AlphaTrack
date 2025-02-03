package util

import (
	"context"
	"fmt"
	"time"
)

func DoWithTimeout(duration time.Duration, task func() error) error {
	ctx, cancel := context.WithTimeout(context.Background(), duration)
	defer cancel()
	
	done := make(chan error)
	
	go func() {
			done <- task()
	}()

	select {
	case err := <-done:
			return err
	case <-ctx.Done():
			return fmt.Errorf("timeout: %w", ctx.Err())
	}
}