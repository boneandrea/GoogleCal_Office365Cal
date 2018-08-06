test:
	bundle exec rspec spec/*

lint:
	bundle exec rubocop --auto-correct *rb
