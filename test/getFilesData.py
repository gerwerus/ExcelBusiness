import os
def getFilesData(planPath="plans"):
    data = []
    for top, _, files in os.walk(planPath):
        for file in files:
            data.append(os.path.join(top, file))
    return data
print(getFilesData("plans"))