.PHONY: test repl fmt clean

# Canonical test command (Roswell, clean session)
test:
	XDG_CACHE_HOME=$(PWD)/build ros -Q scripts/test.ros

# Optional: start an interactive REPL that loads the local .asd and quickloads the system
repl:
	ros -Q run --load scripts/repl.ros

# Placeholder targets (keep for future tooling)
fmt:
	@echo "No formatter configured yet."

clean:
	@echo "Nothing to clean yet."
