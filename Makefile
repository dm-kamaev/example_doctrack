build:
	docker build -t example_doctrack --no-cache .;

remove:
	docker image rm -f example_doctrack;

up:
	docker run -t example_doctrack;

run_and_save_to_file:
	rm -f result.txt;
	docker run -it example_doctrack sh -c "TRACKING_PIXEL_URL=/img/2342f/ node index.js" > result.txt;


doctrack_and_save_to_file:
	rm -f output/result.docx;
	# docker run --name example_doctrack -v /Start/example_doctrack/output:/app/output example_doctrack;
	# docker run \
	# 	-v /Start/example_doctrack/output:/app/output \
	# 	-it example_doctrack sh -c "dotnet ./doctrack/bin/Release/net6.0/linux-x64/doctrack.dll -i doctrack_template.docx -o output/doctrack_example.docx --metadata metadata.json --url http://localhost:5001/image.png;";
	docker run \
		-v /Start/example_doctrack/output:/app/output \
		-v /Start/example_doctrack/input:/app/input \
		--network=cc0 \
		-it example_doctrack sh -c \
		"./doctrack -i input/test.docx -o output/output.docx --url http://localhost:5001/image.png;";
	sudo chown -R dkamaev:dkamaev output;
	sudo chmod -R 0777 output;

  # "mc alias set myminio http://minio:10080 deception deception; mc cp myminio/container-logs/doctrack_template.docx /app/input; ./doctrack -i input/doctrack_template.docx -o output/doctrack_example.docx --metadata metadata.json --url http://localhost:5001/image.png; mc cp /app/output/doctrack_example.docx myminio/container-logs/;";
	# -it example_doctrack sh -c "mc alias set myminio http://minio:10080 deception deception;mc admin info myminio;mc cp myminio/container-logs/logs_localhost_20240304_010650.tar.gz /app/output";
	# -it example_doctrack sh -c "mc alias set myminio http://minio:10080 deception deception;mc admin info myminio";
	# -it example_doctrack sh -c "./doctrack -i input/doctrack_template.docx -o output/doctrack_example.docx --metadata metadata.json --url http://localhost:5001/image.png;";


# docker run -it example_doctrack sh -c "ls -lah doctrack/bin/Release/net6.0/linux-x64/doctrack.dll --help";
# docker run -it example_doctrack sh -c "dotnet ./doctrack/bin/Release/net6.0/linux-x64/doctrack.dll -i doctrack_template.docx -o doctrack_example.docx --metadata metadata.json --url http://test.url/image.png; ls -lah";
# rm -f result.txt;

# Example with input and output option
# docker run \
# 	-v /Start/example_doctrack/output:/app/output \
# 	-it example_doctrack sh -c "dotnet ./doctrack/bin/Release/net6.0/linux-x64/doctrack.dll -i doctrack_template.docx -o output/doctrack_example.docx --metadata metadata.json --url http://localhost:5001/image.png;";
