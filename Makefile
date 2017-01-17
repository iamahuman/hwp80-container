PREFIX=i686-w64-mingw32-
CXX=$(PREFIX)g++
WINDRES=$(PREFIX)windres
CXXFLAGS=-Wall -ggdb -O1
LDFLAGS=-Wall -nostdlib -lkernel32 -luser32 -lole32 -luuid -loleaut32 -lcomctl32 -lshell32 -lshlwapi -Wl,--subsystem,windows
TARGETW=containerw.exe
TARGETC=container.exe
RUNNER=wine

all: $(TARGETW) $(TARGETC)

clean:
	rm -f $(TARGET) *.o *.res

run: $(TARGETC)
	$(RUNNER) $^

$(TARGETW): main.o utils.o main.res
	$(CXX) -o $@ $^ -Wl,-e_start $(LDFLAGS) -Wl,--subsystem,windows

$(TARGETC): main.o utils.o main.res
	$(CXX) -o $@ $^ -Wl,-e_start $(LDFLAGS) -Wl,--subsystem,console

main.o: main.cpp hwpctrl.h common.h
	$(CXX) -c -o $@ $< $(CXXFLAGS)

utils.o: utils.cpp common.h
	$(CXX) -c -o $@ $< $(CXXFLAGS)

main.res: main.rc container.exe.manifest
	$(WINDRES) --input $< --output $@ --output-format=coff

.PHONY: all clean run
