COMPOSE = docker compose
SERVICE = app
GAP_M  ?= 30

all: run

start-docker:
	sudo systemctl start docker

build:
	$(COMPOSE) build

run: build
	$(COMPOSE) run --rm $(SERVICE)

down:
	$(COMPOSE) down --remove-orphans

clean: down
	$(COMPOSE) down --rmi all

fclean: clean
	rm -rf data

re: fclean all

.PHONY: all build run down clean fclean re start-docker